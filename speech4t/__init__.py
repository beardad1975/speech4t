import win32com.client
import speech_recognition as sr
import threading
import time

from . import common

__all__ = [ 
            '語音合成', '設定語音音量', '設定語音速度', '語音說完了嗎',
            '語音辨識google', '辨識成功了嗎', '取得辨識文字',
            '等待語音說完','語音辨識azure', '暫停語音辨識',
            '繼續語音辨識', '語音辨識中嗎',
            ]

# tts init
common.speaker = win32com.client.Dispatch("SAPI.SpVoice")

common.speaker.Volume = common.DEFAULT_VOLUME
common.speaker.Rate = common.DEFAULT_RATE

# recognization init
common.recognizer = sr.Recognizer()
common.recognizer.pause_threshold = 0.5
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

def 語音說完了嗎(ms=1):
    return common.speaker.WaitUntilDone(ms)

def 等待語音說完():
    return common.speaker.WaitUntilDone(-1)

#### recog wrapper function
def recog_callback(recognizer, audio):
    if common.recog_paused :
        print('pausing')
        return

    try:
        if common.recog_service == 'google':
            text = recognizer.recognize_google(audio,language="zh-TW" )
        elif common.recog_service == 'azure':
            text = recognizer.recognize_azure(audio,language="zh-TW",
                    key=common.recog_key, location=common.recog_location )

        if text :
            print('<<',common.recog_service, '辨識為: ', text,'>>')
            
            with common.lock:
                common.recog_text = text
            

            common.recog_countdown -= 1
            if common.recog_countdown <= 0 :
                common.stopper(wait_for_stop=False)
                common.recog_service = False
                print('<<超過次數，語音辨識程式停止>>')

    except sr.UnknownValueError:
            print("<<語音內容無法辨識>>")
            common.recog_countdown -= 1
    except sr.RequestError as e:
            print('<<',common.recog_service,"語音辦識無回應(可能無網路或是超過限制)>>: {0}".format(e))
            common.recog_countdown -= 1

########## rewrite background listen recog thread
def recog_thread():
    with common.mic as source:
        print('<<校正麥克風...>>')
        common.recognizer.adjust_for_ambient_noise(source)  

        while True:
            if common.recog_paused:
                #print('pausing')
                time.sleep(0.1)
                continue

            voice = common.recognizer.listen(source, phrase_time_limit=5)

            if common.recog_discard:
                #print('paused during listening, discard voice')
                with common.lock:
                    common.recog_discard = False
                continue 

            try:
                if common.recog_service == 'google':
                    text = common.recognizer.recognize_google(voice,language="zh-TW" )
                elif common.recog_service == 'azure':
                    text = common.recognizer.recognize_azure(voice,language="zh-TW",
                        key=common.recog_key, location=common.recog_location )

                #text = common.recognizer.recognize_google(voice, language="zh-tw")
                print("<<Google 語音辨識為:", text,'>>')

                if common.recog_discard:
                    #print('paused during recognizing, discard text')
                    with common.lock:
                        common.recog_discard = False
                    continue

                if text:
                    with common.lock:
                        common.recog_text = text
                    common.recog_countdown -= 1
                    if common.recog_countdown <= 0 :
                        print('<<超過次數，語音辨識程式停止>>')
                        common.recog_service = False
                        break
                        
            except sr.UnknownValueError:
                    print("<<語音內容無法辨識>>")
                    common.recog_countdown -= 1
            except sr.RequestError as e:
                    print('<<',common.recog_service,"語音辦識無回應(可能無網路或是超過限制)>>: {0}".format(e))
                    common.recog_countdown -= 1       


def 語音辨識google(次數=15):
    if common.recog_service:
        print("<<語音辨識已啟動>>")
        return 

    # start recog service
    common.recog_countdown = 次數
    common.recog_service = 'google'
    common.recog_paused = False

    t = threading.Thread(target=recog_thread)
    t.daemon = True
    t.start()  

    print('<<開始語音辨識: 採google服務>>\n<<請說話>>')


# def 語音辨識google(次數=15):
#     if common.recog_service:
#         print("<<語音辨識已啟動>>")
#         return 

#     # start recog service
#     with common.mic as source:
#         print('<<校正麥克風...>>')
#         common.recognizer.adjust_for_ambient_noise(source)    
#     common.stopper = common.recognizer.listen_in_background(
#                 common.mic, recog_callback, phrase_time_limit=10)
#     print('<<開始語音辨識: 採google服務>>\n<<請說話>>')
#     common.recog_countdown = 次數
#     common.recog_service = 'google'
#     common.recog_paused = False

def 語音辨識azure(key, location='westus'):
    if common.recog_service:
        print("<<語音辨識已啟動>>")
        return 

    common.recog_countdown = 1000
    common.recog_service = 'azure'
    common.recog_key = key
    common.recog_location = location
    common.recog_paused = False

    t = threading.Thread(target=recog_thread)
    t.daemon = True
    t.start() 

    print('<<開始語音辨識: 採azure服務>>\n<<請說話>>')

# def 語音辨識azure(key, location='westus'):
#     if common.recog_service:
#         print("<<語音辨識已啟動>>")
#         return 

#     with common.mic as source:
#         print('<<校正麥克風...>>')
#         common.recognizer.adjust_for_ambient_noise(source)    
#     common.stopper = common.recognizer.listen_in_background(
#                 common.mic, recog_callback, phrase_time_limit=10)
#     print('<<開始語音辨識: 採azure服務>>\n<<請說話>>')
#     common.recog_countdown = 1000
#     common.recog_service = 'azure'
#     common.recog_key = key
#     common.recog_location = location
#     common.recog_paused = False

# def 關閉語音辨識():
#     if common.recog_service:
#         common.stopper(wait_for_stop=False)
#         common.recog_service = False
#         print('<<語音辨識程式停止>>')
#     else:
#         print('<<無語音辨識程式>>')        


def 辨識成功了嗎():
    if common.recog_service and common.recog_text:
        return True
    else:
        return False
    
def 取得辨識文字():
    tmp = common.recog_text
    with common.lock:
        common.recog_text = ''
    return tmp

def 暫停語音辨識():
    if not common.recog_paused: 
        with common.lock:
            common.recog_text = ''
            common.recog_paused = True
            common.recog_discard = True
        print('<<語音辨識暫停>>')    

def 繼續語音辨識():
    if common.recog_paused:
        with common.lock:
            common.recog_paused = False
        print('<<語音辨識繼續>>')

def 語音辨識中嗎():
    return  not common.recog_paused and not common.recog_discard


if __name__ == '__main__' :
    pass
    
