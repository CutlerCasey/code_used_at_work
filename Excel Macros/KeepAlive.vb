public declare PtrSafe Sub Sleep Lib "Kernel32" (ByVal Milliseconds As LongPtr)
Public Declare sub mouse_event lib "user32" (byval dwFlags as Long, byVal dx as long, byval cButtons as long, byval dwExtraInfo as Long)
public const mouseeventf_leftdown = &h2
public const mouseeventf_leftup = &h4
public const mouseeventf_rightdown as long = &h8
public const mouseeventf_rightup as long = &h10

declare function getcursorpos lib "user32" (lpPoint as pointapi) as long
declare function setcursorpos lib "user32" (byval x as long, byval y as long) as long

type pointapi
	x_pos as long
	y_pos as long
end type

public const sheetname = "sheet2"

public sub SingleClick()
	dim x as string
	if thisworkook.sheets(sheetname).range("a1") = "running" then
		dim hold as pointapi
		
		dim xRange, yRange as integer
		xRange = RndInt(-10, 10)
		yRange = RndInt(-10, 10)
		while xRange = yRange
			xRange = RndInt(-10, 10)
			yRange = RndInt(-10, 10)
		wend
		
		getcursorpos hold
		setcursorpos hold.x_pos + xrange, hold.y_pos + yrange
		application.wait dateadd("s", RndDbl(.25, .5), now)
		mouse_event mouseeventf_leftdown, 0, 0, 0, 0
		mouse_event mouseeventf_leftup, 0, 0, 0, 0
		application.wait dateadd("s", RndDbl(.25, .5), now)
		setcursorpos hold.x_pos, hold.y_pos

		call SleeperFunk
	else
		call StopAllMacro
	end if
end sub

public sub SleeperFunk()
	thisworkbook.sheets(sheetname).range("a1") = "running"
	thisworkbook.sheets(sheetname).range("a2") = now()
	application.ontime now + timevalue("00:00:" + RndStr(10, 30)), "singleclick"
end sub

public sub StopAllMacro()
	thisworkbook.sheets(sheetname).range("a1") = "stop autoclick"
	thisworkbook.sheets(sheetname).range("a2") = ""
	thisworkbook.sheets(sheetname).range("c1") = "running"
end sub

function RndDbl(lowerBound as double, upperBound as double) as double
	randomize
	rnddbl = lowerBound + rnd() * (upperBound - lowerBound)
end function

function RndInt(lowerBound as integer, upperBound as integer) as Integer
	randomize
	RndInt = int(lowerBound + rnd() * (upperBound - lowerBound + 1))
end function

functio RndStr(lowerBound as integer, upperBound as integer) as string
	RndStr = RndInt(lowerBound, upperBound)
end function
