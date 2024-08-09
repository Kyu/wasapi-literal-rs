use windows::core::GUID;
use windows::Win32::Foundation::RPC_E_CHANGED_MODE;
use windows::Win32::Media::Audio::{AUDCLNT_BUFFERFLAGS_SILENT, AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_LOOPBACK, eConsole, eRender, IAudioCaptureClient, IAudioClient, IAudioRenderClient, IMMDevice, IMMDeviceEnumerator, WAVEFORMATEX};
use windows::Win32::System::Com::{CLSCTX_ALL, CoCreateInstance, COINIT_APARTMENTTHREADED, CoInitializeEx, CoTaskMemFree};
use windows::Win32::System::Performance::{QueryPerformanceCounter, QueryPerformanceFrequency};
use crate::COM_INITIALIZED;

const CLSID_MMDEVICE_ENUMERATOR: GUID = GUID {
    data1: 0xbcde0395,
    data2: 0xe52f,
    data3: 0x467c,
    data4: [ 0x8e, 0x3d, 0xc4, 0x57, 0x92, 0x91, 0x69, 0x2e ]
};
const IID_IMMDEVICE_ENUMERATOR: GUID = GUID {
    data1: 0xa95664d2,
    data2: 0x9614,
    data3: 0x4f35,
    data4: [ 0xa7, 0x46, 0xde, 0x8d, 0xb6, 0x36, 0x17, 0xe6 ]
};
const IID_IAUDIO_CLIENT: GUID = GUID {
    data1: 0x1cb9ad4c,
    data2: 0xdbfa,
    data3: 0x4c32,
    data4: [ 0xb1, 0x78, 0xc2, 0xf5, 0x68, 0xa7, 0x03, 0xb2 ]
};
const IID_IAUDIO_CAPTURE_CLIENT: GUID = GUID {
    data1: 0xc8adbd64,
    data2: 0xe71e,
    data3: 0x48a0,
    data4: [ 0xa4, 0xde, 0x18, 0x5c, 0x39, 0x5c, 0xd3, 0x17 ]
};
const IID_IAUDIO_RENDER_CLIENT: GUID = GUID {
    data1: 0xf294acfc,
    data2: 0x3146,
    data3: 0x4483,
    data4: [ 0xa7, 0xbf, 0xad, 0xdc, 0xa7, 0xc2, 0x60, 0xe2 ]
};

// from cpal
// Use RAII to make sure CoTaskMemFree is called when we are responsible for freeing.
struct WaveFormatExPtr(*mut WAVEFORMATEX);

impl Drop for WaveFormatExPtr {
    fn drop(&mut self) {
        unsafe {
            CoTaskMemFree(Some(self.0 as *mut _));
        }
    }
}

struct AudioCaptureData {
    samples: Vec<u64>,
    count: usize,
    time: u64
}

pub struct AudioCapture {
    play_client: Option<IAudioClient>,
    capture_client: Option<IAudioClient>,
    capture: Option<IAudioCaptureClient>,
    format: Option<*mut WAVEFORMATEX>,
    start_qpc: i64,
    start_pos: i64,
    freq: i64,
    use_device_timestamp: bool,
    first_time: bool
}

impl AudioCapture {
    pub(crate) fn new() -> Self {
        COM_INITIALIZED.with(|_| {});

        Self {
            play_client: None,
            capture_client: None,
            capture: None,
            format: None,
            start_qpc: 0,
            start_pos: 0,
            freq: 0,
            use_device_timestamp: false,
            first_time: false,
        }
    }

    // CoInitializeEx must always be called first!
    pub(crate) unsafe fn start(&mut self, duration_100ns: i64) -> bool {
        println!("Starting!");
        let result = false;

        // HR(CoCreateInstance(&CLSID_MMDEVICE_ENUMERATOR, NULL, CLSCTX_ALL, &IID_IMMDEVICE_ENUMERATOR, (LPVOID*)&enumerator));
        let enumerator: IMMDeviceEnumerator = CoCreateInstance(&CLSID_MMDEVICE_ENUMERATOR, None, CLSCTX_ALL).unwrap();
        // CoCreateInstance::<_, Audio::IMMDeviceEnumerator>

        // if (FAILED(IMMDeviceEnumerator_GetDefaultAudioEndpoint(enumerator, eRender, eConsole, &device)))
        // do nothing
        match enumerator.GetDefaultAudioEndpoint(eRender, eConsole) {
            Err(e) => println!("Error in GetDefaultAudioEndpoint: {}", e), // TODO error! logging
            Ok::<IMMDevice, _>(device) => {
                // setup playback for silence, otherwise loopback recording does not provide any data if nothing is playing
                {
                    let length: i64 = 10 * 1000 * 1000; // Original type REFERENCE_TIME :: LONGLONG


                    // can fail if the device has been disconnected since we enumerated it, or if
                    // the device doesn't support playback for some reason
                    let client: IAudioClient = device.Activate(CLSCTX_ALL, None).unwrap();
                    // HR(IMMDevice_Activate(device, &IID_IAUDIO_CLIENT, CLSCTX_ALL, NULL, (LPVOID*)&client));

                    let format = client.GetMixFormat().map(WaveFormatExPtr).unwrap().0;
                    // HR(IAudioClient_GetMixFormat(client, &format));

                    client.Initialize(AUDCLNT_SHAREMODE_SHARED, 0, length, 0, format, None).expect("playback client.Initialize failed");
                    // HR(IAudioClient_Initialize(client, AUDCLNT_SHAREMODE_SHARED, 0, length, 0, format, NULL));

                    let render: IAudioRenderClient = client.GetService().unwrap();
                    // HR(IAudioClient_GetService(client, &IID_IAUDIO_RENDER_CLIENT, &render));

                    let buffer: *mut u8 = render.GetBuffer((*format).nSamplesPerSec).unwrap();
                    // HR(IAudioRenderClient_GetBuffer(render, format->nSamplesPerSec, &buffer));

                    render.ReleaseBuffer((*format).nSamplesPerSec, AUDCLNT_BUFFERFLAGS_SILENT.0 as u32).expect("playback render.ReleaseBuffer failed");
                    // HR(IAudioRenderClient_ReleaseBuffer(render, format->nSamplesPerSec, AUDCLNT_BUFFERFLAGS_SILENT));

                    // IAudioRenderClient_Release(render);

                    // CoTaskMemFree(format);

                    client.Start().expect("playback client.Start failed");
                    // HR(IAudioClient_Start(client));

                    self.play_client = Some(client);
                }

                // loopback recording, TODO: mic recording
                {
                    let client: IAudioClient = device.Activate(CLSCTX_ALL, None).unwrap();
                    // HR(IMMDevice_Activate(device, &IID_IAUDIO_CLIENT, CLSCTX_ALL, NULL, (LPVOID*)&client));

                    let format = client.GetMixFormat().map(WaveFormatExPtr).unwrap().0;
                    // HR(IAudioClient_GetMixFormat(client, &format));

                    client.Initialize(AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_LOOPBACK, duration_100ns, 0, format, Some(&IID_IAUDIO_CAPTURE_CLIENT)).expect("loopback client.Initialize failed");
                    // HR(IAudioClient_Initialize(client, AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_LOOPBACK, duration_100ns, 0, format, NULL));
                    
                    client.GetService::<IAudioCaptureClient>().expect("loopback client.GetService failed");
                    // HR(IAudioClient_GetService(client, &IID_IAUDIO_CAPTURE_CLIENT, (LPVOID*)&capture->capture));

                    client.Start().expect("loopback client.Start failed");
                    // HR(IAudioClient_Start(client));


                    // TODO ??
                    let start: *mut i64 = &mut 0i64;
                    // LARGE_INTEGER start;

                    QueryPerformanceCounter(start).expect("loopback QPC(*start) failed");
                    // QueryPerformanceCounter(&start);
                    
                    self.start_qpc = *start; //.quad_part;

                    let freq: *mut i64 = &mut 0i64;
                    // LARGE_INTEGER freq;

                    QueryPerformanceFrequency(freq).expect("loopback");
                    // QueryPerformanceFrequency(&freq);

                    self.freq = *freq; //.quad_part;


                    self.capture_client = Some(client);
                    self.format = Some(format);
                    self.start_pos = 0;
                    self.use_device_timestamp = true;
                    self.first_time = true;
                }
            }
        }

        // mmdeviceapi.h { IMMDeviceEnumerator_Release }
        // IMMDeviceEnumerator_Release(enumerator);
        result
    }

    pub(crate) unsafe fn get_buffer_frame_count(&mut self) {

    }

    pub(crate) unsafe fn stop(self) {
        println!("Stopping!");
        CoTaskMemFree(Some(self.format.unwrap() as *mut _));
        // self.capture.unwrap().ReleaseBuffer(0).expect("capture.ReleaseBuffer failed");
        self.capture_client.unwrap().Stop().expect("TODO2");
        self.play_client.unwrap().Stop().expect("TODO3");
    }

    fn flush() {}

    fn get_data() {}

    fn release_data() {}
}