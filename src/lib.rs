use windows::Win32::Foundation::RPC_E_CHANGED_MODE;
use windows::Win32::System::Com::{COINIT_APARTMENTTHREADED, CoInitializeEx};

mod audio_capture;

thread_local!(static COM_INITIALIZED: bool = unsafe {
    let result = CoInitializeEx(None, COINIT_APARTMENTTHREADED);

    if result.is_ok() || result == RPC_E_CHANGED_MODE {
        true
    } else {
        panic!(
            "Failed to initialize COM: {}",
            result.0
        );
    }
});

pub fn add(left: u64, right: u64) -> u64 {
    left + right
}

#[cfg(test)]
mod tests {
    use std::thread::{sleep, Thread};
    use std::time::Duration;
    use super::*;


    #[test]
    fn two_plus_two() {
        let result = add(2, 2);
        assert_eq!(result, 4);
    }

    #[test]
    fn audio_cap() {
        let mut x = audio_capture::AudioCapture::new();
        unsafe {
            x.start(1_000_000_000);
            sleep(Duration::from_secs(5));
            x.stop();
        }
    }
}
