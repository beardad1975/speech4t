import win32com.client

from . import common

__all__ = [ 
            '語音合成', '設定語音音量', '設定語音速度', '語音說完了嗎',
            ]

# tts
common.speaker = win32com.client.Dispatch("SAPI.SpVoice")

common.speaker.Volume = common.DEFAULT_VOLUME
common.speaker.Rate = common.DEFAULT_RATE

### Custom Exceptions
# class ImageReadError(Exception):
#     def __init__(self, value):
#         message = f"無法讀取圖片檔 (檔名:{value})"
#         super().__init__(message)

# stt



### wrapper functions

def 語音合成(text, 等待=True):
    if 等待:
        common.speaker.Speak(text, common.SVSFDefault)
    else:
        common.speaker.Speak(text, common.SVSFlagsAsync)

def 設定語音音量(volume):
    volume =  max(min(volume,100), 0)
    common.speaker.Volume = volume

def 設定語音速度(rate):
    rate = max(min(rate,10), -10)
    common.speaker.Rate = rate

def 語音說完了嗎(ms=500):
    return common.speaker.WaitUntilDone(ms)


if __name__ == '__main__' :
    pass
    
