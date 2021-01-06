import win32com.client
import speech_recognition as sr
import threading

from . import common

__all__ = [ 
            '語音合成', '設定語音音量', '設定語音速度', '語音說完了嗎',
            '語音辨識google', '辨識成功了嗎', '取得辨識文字',
            '等待語音說完','語音辨識azure',
            ]

# tts init
common.speaker = win32com.client.Dispatch("SAPI.SpVoice")

common.speaker.Volume = common.DEFAULT_VOLUME
common.speaker.Rate = common.DEFAULT_RATE

# recognization init
common.recognizer = sr.Recognizer()
common.mic = sr.Microphone()
common.lock = threading.Lock()


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

def 語音說完了嗎(ms=100):
    return common.speaker.WaitUntilDone(ms)

def 等待語音說完():
    return common.speaker.WaitUntilDone(-1)

#### recog wrapper function
def recog_callback(recognizer, audio):
    try:
        if common.recog_service == 'google':
            text = recognizer.recognize_google(audio,language="zh-TW" )
        elif common.recog_service == 'azure':
            text = recognizer.recognize_azure(audio,language="zh-TW",
                    key=common.recog_key, location=common.recog_location )

        if text :
            print(common.recog_service, '辨識為: ', text)
            
            with common.lock:
                common.recog_text = text
            

            common.recog_countdown -= 1
            if common.recog_countdown <= 0 :
                common.stopper(wait_for_stop=False)
                common.recog_service = False
                print('<<超過次數，辨識程式停止>>')

    except sr.UnknownValueError:
            print("語音內容無法辨識")
            common.recog_countdown -= 1
    except sr.RequestError as e:
            print(common.recog_service,"語音辦識無回應(可能無網路或是超過限制): {0}".format(e))
            common.recog_countdown -= 1


def 語音辨識google(次數=10):
    with common.mic as source:
        print('校正麥克風...')
        common.recognizer.adjust_for_ambient_noise(source)    
    common.stopper = common.recognizer.listen_in_background(
                common.mic, recog_callback, phrase_time_limit=10)
    print('開始語音辨識: 採google服務\n請說話')
    common.recog_countdown = 次數
    common.recog_service = 'google'

def 語音辨識azure(key, location='westus'):
    with common.mic as source:
        print('校正麥克風...')
        common.recognizer.adjust_for_ambient_noise(source)    
    common.stopper = common.recognizer.listen_in_background(
                common.mic, recog_callback, phrase_time_limit=10)
    print('開始語音辨識: 採azure服務\n請說話')
    common.recog_countdown = 1000
    common.recog_service = 'azure'
    common.recog_key = key
    common.recog_location = location




def 辨識成功了嗎():
    return True if common.recog_text else False
    
def 取得辨識文字():
    tmp = common.recog_text
    with common.lock:
        common.recog_text = ''
    return tmp

if __name__ == '__main__' :
    pass
    
