from win32com.client import Dispatch
import json
import os
import wave
import vosk
import pyaudio
import numpy as np
import sys,time
import requests
from bs4 import BeautifulSoup
import subprocess
from lxml import etree
from openai import OpenAI
os.system('cls') #清屏
speaker = Dispatch("SAPI.SpVoice")#初始化语音



while True:
    tuichuxunhuan = input("\033[94m回车继续 \033[0m")

    if tuichuxunhuan == '':
        print("\033[94m-------------------------------------\033[0m")
        def print_one_by_one(text):        

                sys.stdout.write("\r " + " " * 60 + "\r")    #/r 光标回到行首, \n 换行

                sys.stdout.flush() # 把缓冲区全部输出

                for c in text:

                    sys.stdout.write(c)
                    sys.stdout.flush()
                    time.sleep(0.08)
        def speech_to_text():
            global result
            CHUNK = 1024  # 录音参数配置，指定每次读取音频数据的大小
            FORMAT = pyaudio.paInt16  # 音频格式
            CHANNELS = 1  # 声道数
            RATE = 16000  # 采样率
            p = pyaudio.PyAudio()  # 创建PyAudio对象
            stream = p.open(format=FORMAT, channels=CHANNELS, rate=RATE, input=True, frames_per_buffer=CHUNK)  # 打开音频输入流
            frames = []  # 存储音频数据
            silent_count = 0  # 静音计数器，用于判断是否结束录音
            print("\033[94m语音助手启动\033[0m")
            while True:  # 持续录音，直到检测到2秒钟内没有声音
                data = stream.read(CHUNK)  # 从音频输入流中读取数据
                frames.append(data)  # 将数据添加到frames列表中
                audio_data = np.frombuffer(data, dtype=np.int16)  # 将二进制数据转换为numpy数组
                rms = np.sqrt(np.mean(np.square(audio_data)))  # 计算音量大小（RMS）
                if rms < 4000:  # 判断是否静音
                    silent_count += 1
                else:
                    silent_count = 0
                if silent_count > 2 * RATE / CHUNK:  # 如果持续2秒钟（40分贝以下），则停止录音
                    break
            stream.stop_stream()  # 关闭音频输入流
            stream.close()
            p.terminate()  # 关闭PyAudio对象
            filename = "cache/record.wav"  # 保存录音文件的路径和文件名
            wf = wave.open(filename, 'wb')  # 创建一个Wave_write对象，用于写入音频数据
            wf.setnchannels(CHANNELS)  # 设置声道数
            wf.setsampwidth(p.get_sample_size(FORMAT))  # 设置采样宽度
            wf.setframerate(RATE)  # 设置采样率
            wf.writeframes(b''.join(frames))  # 将音频数据写入文件
            wf.close()  # 关闭文件
            model_path = "model/vosk-model-small-cn-0.22"  # vosk模型文件的路径
            sample_rate = 16000  # 样本采样率
            if not os.path.exists(model_path):  # 检查模型文件是否存在
                print(f"模型路径 {model_path} 不存在，请确保已下载正确的模型文件.")
                return
            model = vosk.Model(model_path)  # 加载vosk模型
            recognizer = vosk.KaldiRecognizer(model, sample_rate)  # 创建识别器对象
            wf = wave.open(filename, 'rb')  # 打开录音文件
            if wf.getnchannels() != 1 or wf.getsampwidth() != 2 or wf.getcomptype() != "NONE":  # 检查录音文件格式是否符合要求
                print("录音文件格式不符合要求.")
                return
            recognizer.SetWords(True)  # 设置识别器输出结果中包含单词信息
            while True:
                data = wf.readframes(4000)  # 一次读取4000个字节的音频数据
                if len(data) == 0:  # 音频数据已全部读取完毕
                    break
                if recognizer.AcceptWaveform(data):  # 将音频数据传入识别器进行识别
                    result = recognizer.Result()  # 获取识别结果
            try:
                result_json = json.loads(result)  # 将字符串转换为JSON对象
                result_text = result_json["text"].replace(" ", "")  # 提取"text"中的文字并去掉空格
            except:
                result_text = ""
                print("录音转文字失败，请重试")
            return result_text


        result = speech_to_text()
        print("主人:", result)



        # 检测前两个字符是否是“搜索”
        if result[:2] == "搜索":
            # 如果是，则将搜索后面的全部字符赋值给另一个变量
            search_content = result[2:]
            print("即将自动联网搜索内容:", search_content)
            



            def search_bing(query):
                url = "https://www.bing.com/search?q=" + query
                headers = {'User-Agent': 'Mozilla/5.0'}
                response = requests.get(url, headers=headers)

                soup = BeautifulSoup(response.text, 'html.parser')
                results = soup.find_all('li', class_='b_algo')

                for i, result in enumerate(results, start=1):
                    title = result.find('h2').text
                    link = result.find('a')['href']
                    snippet = result.find('p').text
                    print_one_by_one(f"{i}. {title}")
                    print(f"\033[94m   {link}\033[0m")
                    print_one_by_one(f"   {snippet}\n")
                    print("\033[96m --------------------------------------------------------\033[0m \n")
                    speaker.Speak(f"{i}. {title}")
            search_bing(search_content)
            

        elif result[:2] == "记录":
            # 如果是，则将搜索后面的全部字符赋值给另一个变量
            jilu = result[2:]
            def create_and_write_to_file(content, filename='new_file(记录).txt'):
                # 获取桌面路径
                desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
                    
                # 创建文件路径
                file_path = os.path.join(desktop_path, filename)
                    
                # 写入内容到文件
                with open(file_path, 'w') as file:
                    file.write(content)
                    
                # 打开记事本编辑文件
                subprocess.Popen(['notepad.exe', file_path]).wait()
                speaker.Speak("记录完成，已通过记事本打开")

            # 要写入的内容
            content = jilu

            # 调用函数创建文件并写入内容
            create_and_write_to_file(content)
            print_one_by_one("-----记录完成-----\n")
            

        elif result[:4] == "今天天气":
            url = f'https://www.tianqishi.com/weichang.html'
            headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36',
            }

            res_data = requests.get(url=url,headers=headers)
            tree = etree.HTML(res_data.text)
            city = tree.xpath('//h3[@class="city-title ico"]')[0].text
            date = tree.xpath('//h3[@class="city-title ico"]//span')[0].text
            ot = tree.xpath('//div[@class="ltlTemperature"]//b')[0].text # 室外温度
            st = tree.xpath('//div[@class="ltlTemperature"]//span')[0].text # 体感温度
            t_type = tree.xpath('(//div[@class="box pcity"])[3]//li//a[@target="_blank"]')[0].text.split('：')[1].split('，')[0]
            all_day_t = tree.xpath('(//div[@class="box pcity"])[3]//li//a[@target="_blank"]')[0].text.split('：')[1].split('，')[1]
            datas = tree.xpath('//ul[@class="mt"]//li')
            values = tree.xpath('//ul[@class="mt"]//li//span')
            he = tree.xpath('(//div[@class="air-quality pd0"])[1]//font')
            suggest = tree.xpath('(//div[@class="air-quality pd0"])[2]//font')
            tianqijianbao = tree.xpath('//div[@class="jdjianjie"]//p')[0]

            print(f"【城市】{city}\n【日期】{date}\n【室外温度】{ot}\n【体感温度】{st}\n【天气情况】{t_type}\n"
                f"【全天气温】{all_day_t}")
            speaker.Speak(f"【城市】{city}\n【日期】{date}\n【室外温度】{ot}\n【体感温度】{st}\n【天气情况】{t_type}\n"
                f"【全天气温】{all_day_t}")
            for i in range(len(datas)):
                print(f"【{datas[i].text}】{values[i].text}")
            print(f"【健康影响】{he[0].text}\n【建议措施】{suggest[0].text}")
            print(f"【天气简报】{tianqijianbao.text}")

        elif result[:4] == "开启对话":
            print_one_by_one("开始调用chatgpt回答\n--------------------------------------------------\n")
            duihua = result[4:]
            client = OpenAI(
                # defaults to os.environ.get("OPENAI_API_KEY")
                api_key="",
                base_url=""
            )



            # 非流式响应
            def gpt_35_api(messages: list):
                completion = client.chat.completions.create(model="gpt-3.5-turbo", messages=messages)
                print(completion.choices[0].message.content)

            def gpt_35_api_stream(messages: list):
                stream = client.chat.completions.create(
                    model='gpt-3.5-turbo',
                    messages=messages,
                    stream=True,
                )
                for chunk in stream:
                    if chunk.choices[0].delta.content is not None:
                        print(chunk.choices[0].delta.content, end="")

            if __name__ == '__main__':
                messages = [{'role': 'user','content': duihua },]
                gpt_35_api_stream(messages)

        elif result[:4] == "今日新闻":
            
            url = "https://www.toutiao.com/api/pc/feed/?category=news_hot&utm_source=toutiao&widen=1&max_behot_time="
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.100 Safari/537.36"
            }

            def get_news(url):
                response = requests.get(url, headers=headers)
                if response.status_code == 200:
                    data = response.json()
                    for item in data.get("data"):
                        title = item.get("title")
                        abstract = item.get("abstract")
                        source_url = item.get("source_url")
                        time_str = item.get("behot_time")
                        time_stamp = int(time_str)
                        time_array = time.localtime(time_stamp)
                        time_str = time.strftime("%Y-%m-%d %H:%M:%S", time_array)
                        print(f"\033[94m   --------------{title}--------------\033[0m")
                        print(abstract)
                        print(f"\033[96m {time_str}\033[0m \n")
                        speaker.Speak(time_str)
                        speaker.Speak(title)
                        

            if __name__ == '__main__':
                get_news(url)

        else:
            print("未知命令")
            speaker.Speak("未知命令")
    else:
        print("退出语音助手。")
        break

