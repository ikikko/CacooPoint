@Grab(group='org.codehaus.groovy.modules.http-builder', module='http-builder', version='0.5.0')

import builder.PowerPointBuilder
import groovyx.net.http.HTTPBuilder
import static groovyx.net.http.Method.GET
import static groovyx.net.http.ContentType.BINARY

// ============================================================ //
// Cacooの図形IDと（必要ならば）APIキーを指定
def diagramId = 'IZczfsbRj5yEbYEG'
def apiKey = '' // 公開されている図形ならば不要
// ============================================================ //

File imageDir = new File('image')
imageDir.deleteDir()
imageDir.mkdir()

// Cacooの図一覧を取得
def sheetIndex = 1;
def url = "https://cacoo.com/api/v1/diagrams/${diagramId}.xml?apiKey=${apiKey}"
def diagrams = new XmlParser().parse(url)
def title = diagrams.title.text()
diagrams.sheets.sheet.each {
	def name = it.name.text()
	def imageUrl = it.imageUrlForApi.text() + "?apiKey=${apiKey}"
	
	def http = new HTTPBuilder( imageUrl )
	http.request(GET, BINARY) {req ->
		response.success = { resp, reader ->
			def fileName = String.sprintf('%03d_%s.png', sheetIndex, name)
			def sanitizedFileName = fileName.replaceAll(/[\/:*?"<>|]/, ' ')
			new File("image/${sanitizedFileName}") << reader
			sheetIndex++;
			
			println "created image file : ${sanitizedFileName}"
		}
	}
}

// 取得した図をPowerPointに貼りつけ
def builder = new PowerPointBuilder()
builder.slideshow(filename: "${title}.ppt") {
	imageDir.eachFile {
		imageslide(src: it.path)
	}
}
