#	-*- coding:UTF-8 -*-


from xml.dom import minidom
from openpyxl import Workbook, load_workbook
import argparse
import os


#	获取命令行参数
def getCommandLine():
	parser = argparse.ArgumentParser(description = 'get path of the xml file')
	parser.add_argument('xml_path', metavar = 'XML_PATH', type = str, nargs = 1, help = 'path of the xml file')
	
	return parser

# 获取xml文件路径
def getXmlPath():
	parser = getCommandLine()
	args = vars(parser.parse_args())
	path = args['xml_path'][0]

	return path

#	解析xml
def ParseXml(path):

	#	获取DOM对象
	oDom = minidom.parse(path)

	#	获取根节点
	oRoot = oDom.documentElement # <xmeml>

	oVideo = oRoot.getElementsByTagName('video')
	oTracks = oVideo[0].getElementsByTagName('track')

	#	获取第一轨的片段 剪辑轨
	oEditClips = oTracks[0].getElementsByTagName('clipitem')
	#	获取第三轨的片段 back轨
	oVFXClips_Back = oTracks[2].getElementsByTagName('clipitem')
	#	获取第四轨的片段 front轨
	oVFXClips_Front = oTracks[3].getElementsByTagName('clipitem')
	#	获取第五轨 文案轨
	oVFXText = oTracks[4].getElementsByTagName('clipitem')



	#	获取刀数信息
	clipNum = []
	for oClip in oEditClips:
		oStart = oClip.getElementsByTagName('start')[0].childNodes[0].nodeValue
		clipNum.append(int(oStart))

	for oClip in oVFXClips_Back:
		oStart = oClip.getElementsByTagName('start')[0].childNodes[0].nodeValue
		clipNum.append(int(oStart))

	clipNum = list(set(clipNum))	#	去重
	clipNum.sort()	#	排序


	#	获取集数
	oSequence = oRoot.getElementsByTagName('sequence')[0]
	EPName = oSequence.getElementsByTagName('name')[0].childNodes[0].nodeValue[0:4]

	#	获取Front轨的入点信息
	FrontTrackStart = []
	for oClip in oVFXClips_Front:
		oStart = oClip.getElementsByTagName('start')[0].childNodes[0].nodeValue
		FrontTrackStart.append(oStart)

	#	获取Back轨的剪辑效果信息
	EffectDatas = []
	for oClip in oVFXClips_Back:
		mydict = {}

		oStart = oClip.getElementsByTagName('start')[0]

		if oClip.getElementsByTagName('effect'):
			oEffectName = oClip.getElementsByTagName('effect')[0].getElementsByTagName('name')[0].childNodes[0].nodeValue
		else:
			oEffectName = ''

		mydict['start'] = oStart.childNodes[0].nodeValue
		mydict['effectname'] = oEffectName

		EffectDatas.append(mydict)	

	#	获取每条文案的位置和内容
	textDatas = []
	index = 0
	for oClip in oVFXText:

		index = index + 1
		oStart = oClip.getElementsByTagName('start')[0].childNodes[0].nodeValue
		oEnd = oClip.getElementsByTagName('end')[0].childNodes[0].nodeValue
		oText = oClip.getElementsByTagName('effect')[0].getElementsByTagName('name')[0].childNodes[0].nodeValue

		_id = str(index).zfill(4)
		EP = EPName
		clipnum = str(clipNum.index(int(oStart)) + 1).zfill(4)
		start = FrameCountToTimeCode(oStart)
		end = FrameCountToTimeCode(oEnd)
		text = oText
		effectname = ''

		#	是否有剪辑效果
		for item in EffectDatas:
			if item['start'] == str(oStart):
				effectname = item['effectname']							

		#	是否多层
		if oStart in FrontTrackStart:
			level = 'Back'			
			textDatas.append(MyDict(_id, EP, clipnum, start, end, level, effectname, text))
			level = 'Front'
			textDatas.append(MyDict(_id, EP, clipnum, start, end, level, effectname, text))
		else:
			level = ''
			textDatas.append(MyDict(_id, EP, clipnum, start, end, level, effectname, text))





	ExcelOutput(textDatas, EPName)

#	录入一条dict
def MyDict(_id, EP, clipnum, start, end, level, effectname, text):

	if effectname == 'Time Remap':
		effectname = '变速'
	if effectname == 'Basic Motion':
		effectname = '缩放'

	mydict = {}
	mydict['id'] = _id
	mydict['EP'] = EP
	mydict['clipnum'] = clipnum
	mydict['start'] = start
	mydict['end'] = end
	mydict['level'] = level
	mydict['effectname'] = effectname
	mydict['text'] = text
	return mydict

#	输出表格
def ExcelOutput(data, EPName):

	path = PATH_DESKTOP + '/' + EPName + '.xlsx'
	wb = Workbook()			
	ws = wb.active

	#	第一行
	row = ["id", "集数", "第几刀", "入点", "出点", "多层", "剪辑效果", "特效要求"]
	ws.append(row)

	for item in data:

		row = []
		row.append(item['id'])
		row.append(item['EP'])
		row.append(item['clipnum'])
		row.append(item['start'])
		row.append(item['end'])
		row.append(item['level'])
		row.append(item['effectname'])
		row.append(item['text'])

		ws.append(row)

	wb.save(path)
	wb.close()


#	将帧计数转换成时间码
def FrameCountToTimeCode(frameCount):
	frameCount = int(frameCount)
	h = frameCount // (24 * 60 * 60)
	m = frameCount // (24 * 60) - h * 60
	s = frameCount // 24 - h * 60 * 60 - m * 60
	f = frameCount % 24

	return str(h).zfill(2) + ':' + str(m).zfill(2) + ':' + str(s).zfill(2) + ':' + str(f).zfill(2)









if __name__ == '__main__':
	PATH_DESKTOP = os.path.join(os.path.expanduser('~'), "Desktop")
	path = getXmlPath()
	ParseXml(path)

	




