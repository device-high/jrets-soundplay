import os
import sys
import threading
import random

import tkinter
from tkinter import ttk
from tkinter import filedialog

from PIL import Image
from PIL import ImageGrab
import pyocr
from playsound import playsound
import wave

import pandas
import numpy
import time
import re
from collections import deque
import datetime
import math

# Tesseract-OCR初期設定
path='C:\\Users\\path-for-user\\AppData\\Local\\Tesseract-OCR\\'
os.environ['PATH'] = os.environ['PATH'] + path
pyocr.tesseract.TESSERACT_CMD = r'C:\Users\path-for-user\AppData\Local\Tesseract-OCR\tesseract.exe'
tools = pyocr.get_available_tools()
tool = tools[0]

# GUI初期設定

# GUI操作機能定義
def csv_select():
	filetype = [("CSV","*.csv"), ("すべて","*")]
	global file_path
	file_path = tkinter.filedialog.askopenfilename(filetypes = filetype)
	file_name = os.path.basename(file_path)
	csv_box.insert(tkinter.END, file_name)
	global soundcsv
	soundcsv = pandas.read_csv(file_path, header=0)
	soundcsv['next'] = 0
	soundcsv['soon'] = 0
	start_combobox.config(values=list(soundcsv['Station']))

def type_select(event):
	type = tojime_combobox.get()
	selected_tojime = tojime_lists.loc[type]
	global tojime_X, tojime_Y, tojime_threshold
	tojime_X = selected_tojime['X']
	tojime_Y = selected_tojime['Y']
	tojime_threshold = selected_tojime['threshold']

def start_select(event):
	global startstation
	startstation = start_combobox.get()

def start_button_exec():
	playthread = threading.Thread(target=main)
	global playthread_flag
	playthread_flag = 1
	playthread.start()

def stop_button_exec():
	global playthread_flag
	playthread_flag = 0
	time.sleep(1)
	status_info.config(text="停止")
	status_info.update()
	next_info.config(text="")
	next_info.update()
	dist_info.config(text="")
	dist_info.update()
	file_info.config(text="")
	file_info.update()
	tojimestatus_info.config(text="")
	tojimestatus_info.update()

def soon_select(event):
	global soondistance
	soondistance = soon_combobox.get()

# GUI表示定義
gui_window = tkinter.Tk()
gui_window.geometry("240x97")
gui_window.title("jrets-soundplay")

# デバッグ時は常に最上位表示を有効にしたいので、デバッグ時はコメントアウトを外す
# gui_window.attributes("-topmost", True)

# 1行目：CSV選択
csv_label = tkinter.Label(gui_window, text = "CSV：")
csv_label.grid(row = 0, column = 0)

csv_box = tkinter.Entry(width=23)
csv_box.grid(row = 0, column = 1)

csv_button = tkinter.Button(text="参照",command=csv_select)
csv_button.grid(row = 0, column = 2, sticky=tkinter.W)

# 2行目：形式選択（戸閉灯位置・閾値）
tojime_label = tkinter.Label(gui_window, text = "形式：")
tojime_label.grid(row = 1, column = 0)

tojime_lists = pandas.DataFrame(data=numpy.array([[995, 890, 100],[995, 890, 100]]),index=['E233', '211'],columns=['X', 'Y', 'threshold'])
tojime_combobox = ttk.Combobox(gui_window, height=5, width=20, values=list(tojime_lists.index))
tojime_combobox.bind('<<ComboboxSelected>>', type_select)
tojime_combobox.grid(row = 1, column = 1)

# 3行目：始発駅選択
start_label = tkinter.Label(gui_window, text = "始発：")
start_label.grid(row = 2, column = 0)

start_combobox = ttk.Combobox(gui_window, height=5, width=20)
start_combobox.bind('<<ComboboxSelected>>', start_select)
start_combobox.grid(row = 2, column = 1)

# 4行目：開始停止ボタン
start_button = tkinter.Button(text="開始",command=start_button_exec)
start_button.grid(row = 3, column = 0, sticky=tkinter.E)

stop_button = tkinter.Button(text="停止",command=stop_button_exec)
stop_button.grid(row = 3, column = 1, sticky=tkinter.W)

# 接近放送距離選択（2行目と3行目の空きスペース）
soon_label = tkinter.Label(gui_window, text = "接近：")
soon_label.grid(row = 1, column = 2, sticky=tkinter.W)

soon_lists = pandas.DataFrame(data=numpy.array([[850],[750],[650],[550]]),columns=['soonvalue'])
soon_combobox = ttk.Combobox(gui_window, height=5, width=5, values=list(soon_lists.soonvalue))
soon_combobox.bind('<<ComboboxSelected>>', soon_select)
soon_combobox.grid(row = 2, column = 2)

# 5行目以降：ステータス表示（デバッグ用）、初期画面サイズでは画面外とし、表示したいときは手動リサイズする
status_label = tkinter.Label(gui_window, text = "検出情報（デバッグ用）")
status_label.grid(row = 4, column = 0, columnspan = 3)

status_label = tkinter.Label(gui_window, text = "状態：")
status_label.grid(row = 5, column = 0)

status_info = tkinter.Label(gui_window, text = "停止")
status_info.grid(row = 5, column = 1)

next_label = tkinter.Label(gui_window, text = "次駅：")
next_label.grid(row = 6, column = 0)

next_info = tkinter.Label(gui_window, text = "")
next_info.grid(row = 6, column = 1)

dist_label = tkinter.Label(gui_window, text = "距離：")
dist_label.grid(row = 7, column = 0)

dist_info = tkinter.Label(gui_window, text = "")
dist_info.grid(row = 7, column = 1)

file_label = tkinter.Label(gui_window, text = "音声：")
file_label.grid(row = 8, column = 0)

file_info = tkinter.Label(gui_window, text = "")
file_info.grid(row = 8, column = 1)

tojimestatus_label = tkinter.Label(gui_window, text = "戸閉：")
tojimestatus_label.grid(row = 9, column = 0)

tojimestatus_info = tkinter.Label(gui_window, text = "")
tojimestatus_info.grid(row = 9, column = 1)


# メイン処理部
def main():

	# 始発駅チャイム処理
	target = soundcsv[soundcsv['Station'] == startstation]
	dept = datetime.datetime.strptime(target.iloc[-1]['Dept'], '%H:%M:%S')

	chime_file = wave.open(target.iloc[-1]['Chime'], 'rb')
	chime_time = math.floor(chime_file.getnframes() / chime_file.getframerate())
	chime_file.close()

	voice_file = wave.open(target.iloc[-1]['Voice'], 'rb')
	voice_time = math.floor(voice_file.getnframes() / voice_file.getframerate())
	voice_file.close()

	while playthread_flag:

		status_info.config(text="運転開始検知")
		status_info.update()

		img = ImageGrab.grab()
		img_crop = img.crop((1720, 32, 1884, 64))
		img_crop_gray = img_crop.convert("L")
		builder = pyocr.builders.TextBuilder(tesseract_layout=8)
		ocrcurrent = tool.image_to_string(img_crop_gray, lang="eng", builder=builder)

		# 時刻フォーマットの文字列が得られなかったら1秒待ってループの先頭に戻る
		try:
			current = datetime.datetime.strptime(ocrcurrent, '%H:%M:%S')

		except ValueError as error:
			time.sleep(1)
			continue

		# 時刻フォーマットの文字列が得られたら、現在時刻を基準に始発チャイム再生

		status_info.config(text="始発駅発車チャイム再生")
		status_info.update()

		current = datetime.datetime.strptime(ocrcurrent, '%H:%M:%S')
		delta = dept - current

		chime_wait = delta.seconds - chime_time - voice_time - 4
		time.sleep(chime_wait)
		playsound(target.iloc[-1]['Chime'])
		time.sleep(1)
		playsound(target.iloc[-1]['Voice'])

		# 再生が終わったら戸閉灯チェック後、駅間処理へ進む
		while playthread_flag:

			img = ImageGrab.grab()
			img_gray = img.convert("L")
			tojime = img_gray.getpixel((tojime_X,tojime_Y))

			# 戸閉灯が点灯したら2秒待機後に次駅処理へ
			if tojime > tojime_threshold:
				tojimestatus_info.config(text="閉")
				tojimestatus_info.update()

				time.sleep(2)
				break

			# 点灯してなければ0.1～2.5秒待機してループ
			time.sleep(random.uniform(0.1, 2.5))

		break


	# 始発駅発車後は永久ループで次駅を再生し続ける
	while playthread_flag:

		#次駅確認処理
		while playthread_flag:

			status_info.config(text="次駅検知")
			status_info.update()

			# OCR処理（次駅）
			img = ImageGrab.grab()
			img_crop = img.crop((1760, 104, 1884, 144))
			builder = pyocr.builders.TextBuilder(tesseract_layout=6)
			station = tool.image_to_string(img_crop, lang="eng", builder=builder)

			# OCR結果からCSVの該当行を検索
			target = soundcsv[soundcsv['Station'] == station]

			# 検索結果を確認し、ヒットしていたら次駅ループへ
			if len(target.index) == 1:
				next_info.config(text=station)
				next_info.update()
				break

			# ヒットしなかった場合は0.1～2.5秒待ってループ継続
			time.sleep(random.uniform(0.1, 2.5))


		# 対象駅の発車放送が未放送か確認し、未放送なら放送ループへ
		target = soundcsv[soundcsv['Station'] == station]
		if target.iloc[-1]['next'] == 0:
			status_info.config(text="距離検知（発車放送）")
			status_info.update()

			# 過去読み取った距離を記録する配列の初期化
			distancelist = deque([], maxlen=4)

			while playthread_flag:

				# OCR処理（距離）
				img = ImageGrab.grab()
				img_crop = img.crop((1760, 182, 1850, 218))
				img_crop_gray = img_crop.convert("L")
				builder = pyocr.builders.DigitBuilder(tesseract_layout=8)
				distance = tool.image_to_string(img_crop_gray, lang="eng", builder=builder)

				# ゴミを取り除いたOCR文字列が数字であれば、距離判定に
				distnum = re.sub(r"\D", "", distance)
				if distnum.isdecimal():
					distint = int(distnum)

					dist_info.config(text=str(distnum))
					dist_info.update()

					# はじめてこのループに入った時はとりあえず現距離を入れる
					if sum(distancelist) == 0:
						distancelist.append(distint)

					# 過去4回平均と比較して減少率が50%以内であれば読み取り成功とし、過去の読み取り情報に記録した上で再生判定へ
					if sum(distancelist)/len(distancelist) > distint and sum(distancelist)/len(distancelist) < distint*1.5:
						distancelist.append(distint)

						# 発車音声再生距離になったら音声を再生し、再生済を設定後ループ抜け
						if distint < target.iloc[-1]['Distance'].astype(int):
							status_info.config(text="発車放送再生")
							status_info.update()
							file_info.config(text=target.iloc[-1]['Nextfile'])
							file_info.update()

							playsound(target.iloc[-1]['Nextfile'])
							soundcsv.next[soundcsv['Station'] == station] = 1
							break

				# 再生する距離ではなかった場合は0.1～2.5秒待ってループ継続
				time.sleep(random.uniform(0.1, 2.5))


		# 対象駅の接近放送が未放送か確認し、未放送なら放送ループへ
		target = soundcsv[soundcsv['Station'] == station]
		if target.iloc[-1]['soon'] == 0:
			status_info.config(text="距離検知（接近放送）")
			status_info.update()

			# 過去読み取った距離を記録する配列の初期化
			distancelist = deque([], maxlen=4)

			while playthread_flag:

				# OCR処理（距離）
				img = ImageGrab.grab()
				img_crop = img.crop((1760, 182, 1850, 218))
				img_crop_gray = img_crop.convert("L")
				builder = pyocr.builders.DigitBuilder(tesseract_layout=8)
				distance = tool.image_to_string(img_crop_gray, lang="eng", builder=builder)

				# ゴミを取り除いたOCR文字列が数字であれば、距離判定に
				distnum = re.sub(r"\D", "", distance)
				if distnum.isdecimal():
					distint = int(distnum)

					dist_info.config(text=str(distnum))
					dist_info.update()

					# はじめてこのループに入った時はとりあえず現距離を入れる
					if sum(distancelist) == 0:
						distancelist.append(distint)

					# 過去4回平均と比較して減少率が50%以内であれば読み取り成功とし、過去の読み取り情報に記録した上で再生判定へ
					if sum(distancelist)/len(distancelist) > distint and sum(distancelist)/len(distancelist) < distint*1.5:
						distancelist.append(distint)

						# 発車音声再生距離になったら音声を再生し、再生済を設定後ループ抜け
						if distint < soondistance:
							status_info.config(text="接近放送再生")
							status_info.update()
							file_info.config(text=target.iloc[-1]['Soonfile'])
							file_info.update()

							playsound(target.iloc[-1]['Soonfile'])
							soundcsv.soon[soundcsv['Station'] == station] = 1
							break

				# 再生する距離ではなかった場合は0.1～2.5秒待ってループ継続
				time.sleep(random.uniform(0.1, 2.5))


		# 発車放送・次駅放送が共に放送済の場合、発車チャイムループ
		target = soundcsv[soundcsv['Station'] == station]
		if target.iloc[-1]['next'] + target.iloc[-1]['soon'] == 2:

			# 戸閉灯の消灯から次駅到着を確認
			arrived = 1
			while arrived:
				status_info.config(text="到着検知")
				status_info.update()

				img = ImageGrab.grab()
				img_gray = img.convert("L")
				tojime = img_gray.getpixel((tojime_X,tojime_Y))

				# 戸閉灯が消灯したことを確認したら、発車メロディ再生
				if tojime < tojime_threshold:

					# 到着駅の発車メロディ情報取得
					status_info.config(text="途中駅発車チャイム再生")
					status_info.update()

					dept = datetime.datetime.strptime(target.iloc[-1]['Dept'], '%H:%M:%S')

					chime_file = wave.open(target.iloc[-1]['Chime'], 'rb')
					chime_time = math.floor(chime_file.getnframes() / chime_file.getframerate())
					chime_file.close()

					voice_file = wave.open(target.iloc[-1]['Voice'], 'rb')
					voice_time = math.floor(voice_file.getnframes() / voice_file.getframerate())
					voice_file.close()

					while playthread_flag:

						img = ImageGrab.grab()
						img_crop = img.crop((1720, 32, 1884, 64))
						img_crop_gray = img_crop.convert("L")
						builder = pyocr.builders.TextBuilder(tesseract_layout=8)
						ocrcurrent = tool.image_to_string(img_crop_gray, lang="eng", builder=builder)

						# 時刻フォーマットの文字列が得られなかったら1秒待ってループの先頭に戻る
						try:
							current = datetime.datetime.strptime(ocrcurrent, '%H:%M:%S')

						except ValueError as error:
							time.sleep(1)
							continue

						# 時刻フォーマットの文字列が得られたらチャイム再生
						current = datetime.datetime.strptime(ocrcurrent, '%H:%M:%S')

						delta = dept - current

						# 発車時刻まで余裕がある到着の時は到着時刻基準でチャイム再生、余裕がない時はすぐに再生
						if delta.days == 0:
							wait_time = delta.seconds
						else:
							wait_time = 25

						chime_wait = wait_time - chime_time - voice_time - 4

						if chime_wait < 0:
							chime_wait = 0

						time.sleep(chime_wait)
						playsound(target.iloc[-1]['Chime'])
						time.sleep(1)
						playsound(target.iloc[-1]['Voice'])

						# 再生が終わったら、戸閉灯チェック
						while playthread_flag:

							img = ImageGrab.grab()
							img_gray = img.convert("L")
							tojime = img_gray.getpixel((tojime_X,tojime_Y))

							# 戸閉灯が点灯したことを確認したら、発車メロディ再生済フラグを設定し、ループ終了
							if tojime > tojime_threshold:
								arrived = 0
								time.sleep(2)
								break

							# 点灯していない場合は0.1～2.5秒待ってループ
							time.sleep(random.uniform(0.1, 2.5))

						break



		# 全部キャッチできると思うが、念のため何もヒットしなかったら5秒待ってループの先頭に戻る
		time.sleep(5)


# GUI表示
gui_window.mainloop()