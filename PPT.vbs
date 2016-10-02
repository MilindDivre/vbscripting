Const ppLayoutText = 2

 Set objPPT = CreateObject("PowerPoint.Application")
 objPPT.Visible = True

 Set objPresentation = objPPT.Presentations.Add

 'objPresentation.ApplyTemplate _
 '("C:\Program Files\Microsoft Office\Templates\1033\ProjectStatusReport.potx")
for i =1 to 5
 Set objSlide = objPresentation.Slides.Add(i, ppLayoutText)
 objSlide.shapes(2).textFrame.textRange.text="My Slide"
next
 objPresentation.saveas "c:\demo.pptx"
 msgbox "done!!"
 message="PPT Created on c drive by Pratima Tumbagee"
 set sapi = CreateObject("sapi.spvoice")
 sapi.Speak message
 set objPPT = nothing
 
 REM call addSlide("Hello","world")
  REM call addSlide("Hello1","world1")
 REM Function addSlide(title,text)
	REM set objppt = CreateObject("PowerPoint.Application")
	REM objPPT.visible= true
	REM set objPresentation = objPPT.Presentations.add
	REM i=objPresentation.slides.count
	REM i =i+1
	REM Set objSlide = objPresentation.Slides.Add(i, ppLayoutText)
	REM objSlide.shapes(1).textFrame.textRange.text=title
	REM objSlide.shapes(2).textFrame.textRange.text=text
	REM objPresentation.saveas "c:\demo1.pptx"
	REM msgbox "done!!"
	REM set objPresentation = nothing
	REM set objppt = nothing
 REM end function