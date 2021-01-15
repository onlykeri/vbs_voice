Dim i
'speak = inputbox("请设置语音速率，范围（-10，10）","语音播报","1")
speak = 1
Set objVoice = CreateObject("SAPI.SpVoice") 
Set colVoice = objVoice.GetVoices()'获得语音引擎集合
Set objVoice.Voice = colVoice.Item(0)	' 获取语音类型
objVoice.Rate = speak	'调节速度（-10，10）
objVoice.Volume = 80	'调节音量（0-100）
readfilepath=".\六级.txt"
set stm=createobject("ADODB.Stream")
stm.Charset ="utf-8"
stm.Open
stm.LoadFromFile readfilepath
readfile = stm.ReadText 
for i=1 to 3
objVoice.Speak readfile
next
stm.Close