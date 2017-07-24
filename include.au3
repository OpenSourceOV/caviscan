#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         OpenSourceOV.org (http://opensourceov.org)

#ce ----------------------------------------------------------------------------

#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <Date.au3>
#include <Inet.au3>
#include <File.au3>

; Image Types
Const $IMAGE_TYPE_8_GREY = '8 bit Gray'
Const $IMAGE_TYPE_16_GREY = '16 bit Gray'
Const $IMAGE_TYPE_24_RGB = '24 bit RGB'
Const $IMAGE_TYPE_48_RGB = '48 bit RGB'
Const $IMAGE_TYPE_64_RGBI = '64 bit RGBI'

; File Types
Const $FILE_TYPE_TIFF = 1
Const $FILE_TYPE_JPEG = 2

; Misc
Const $GREY_CHANNEL = "Auto" ; which channel to make grey from when using grey image type. Options: Blue, Green, Red, Auto
Const $PREVIEW_RESOLUTION = 75
Const $OUTPUT_FILE_FORMAT = 'YYYY-MM-DD-0001+.jpg'

Global $currentStatus, $currentMode, $currentScanner, $currentFolder, $currentSample, $currentImageType, $currentFileType
Global $finished = False

; Controls
Global $btnPreview, $btnScan
Global $comboCropSize, $comboResolutions, $comboPreviewResolutions, $comboImageTypes, $comboMode, $comboColorBalance, $comboExternalViewer
Global $comboTiffFileType, $comboTiffCompression
Global $comboInfraredClean
Global $comboGreyChannel
Global $statusBar
Global $tabPreviewScan, $tabOptions
Global $radioJpeg, $radioPDF, $radioTIFF
Global $editJpegQuality, $editDefaultFolder, $editFileFormat
Global $editXSize, $editYSize, $editXOffset, $editYOffset

; Log file
Global $logFile = FileOpen($LOG_FILE, 1)

; Scanners
Global $vueScan
Global $scanners[0]
Global $scannerCount = 0

; Events
OnAutoItExitRegister("finish")
; HotKeySet("{ESC}", "die")

wLog('Script started')

; Functions
; ------------------------------------------------------------------------------------
Func scanRepeat($n = 1)
	preScan()
	For $i = 1 To $n
		scanningProcedure()
		Sleep($SCAN_INTERVAL_SEC * 1000)
	Next
	postScan()
	MsgBox($MB_OK, "", "Scans complete")
	FileClose($logFile)
EndFunc

Func scanOnce()
	preScan()
	scanningProcedure()
	postScan()
	MsgBox($MB_OK, "", "Scans complete")
	FileClose($logFile)
EndFunc

Func startScan()

	preScan()

	Do
		if Not(scannerOK()) Then
			fail()
		EndIf

		scanningProcedure()

		Sleep($SCAN_INTERVAL_SEC * 1000)

	Until _DateDiff('s', $END_DATE_TIME, _NowCalc()) > 0

	postScan()

	MsgBox($MB_OK, "", "Scans complete")
	FileClose($logFile)

EndFunc

Func scannerOK()
	$ok = false
	if WinList("[CLASS:wxWindowNR]")[0][0] > 0 Then $ok = true
	If $btnScan Then $ok = true
	Return $ok
EndFunc

Func showAllScanners()
	Local $scannerList = WinList("[CLASS:wxWindowNR]")
	For $i = 1 To $scannerList[0][0]
		WinSetState($scannerList[$i][1], "", @SW_RESTORE)
	Next
EndFunc

Func hideAllScanners()
	Local $scannerList = WinList("[CLASS:wxWindowNR]")
	For $i = 1 To $scannerList[0][0]
		WinSetState($scannerList[$i][1], "", @SW_HIDE)
	Next
EndFunc

Func showScanner($scannerInstance)
	WinSetState($scanners[$scannerInstance - 1], "", @SW_RESTORE)
EndFunc

Func hideScanner($scannerInstance)
	WinSetState($scanners[$scannerInstance - 1], "", @SW_HIDE)
EndFunc

Func hideCurrentScanner()
	hideScanner($scanners[$currentScanner - 1])
EndFunc

Func setCurrentScanner($scannerInstance)

	wLog('Set current scanner to ' & $scannerInstance)

	$currentScanner = $scannerInstance
	$vueScan = $scanners[$scannerInstance - 1]

	setInstanceIDs()

	; Assign Controls
	$btnPreview 				      = ControlGetHandle($vueScan, "", "[CLASS:Button; INSTANCE:1]")
	$btnScan 					        = ControlGetHandle($vueScan, "", "[CLASS:Button; INSTANCE:2]")
	$comboResolutions 			  = ControlGetHandle($vueScan, "", "[CLASS:ComboBox; INSTANCE:64]")
	$comboPreviewResolutions 	= ControlGetHandle($vueScan, "", "[CLASS:ComboBox; INSTANCE:63]")
	$comboImageTypes 			    = ControlGetHandle($vueScan, "", "[CLASS:ComboBox; INSTANCE:60]")
	$comboMode 					      = ControlGetHandle($vueScan, "", "[CLASS:ComboBox; INSTANCE:57]")
	$comboColorBalance			  = ControlGetHandle($vueScan, "", "[CLASS:ComboBox; INSTANCE:19]")
	$comboExternalViewer		  = ControlGetHandle($vueScan, "", "[CLASS:ComboBox; INSTANCE:82]")
	$comboCropSize				    = ControlGetHandle($vueScan, "", "[CLASS:ComboBox; INSTANCE:48]")
	$comboTiffFileType			  = ControlGetHandle($vueScan, "", "[CLASS:ComboBox; INSTANCE:5]")
	$comboTiffCompression		  = ControlGetHandle($vueScan, "", "[CLASS:ComboBox; INSTANCE:6]")
	$comboGreyChannel			    = ControlGetHandle($vueScan, "", "[CLASS:ComboBox; INSTANCE:61]")
	$comboInfraredClean			  = ControlGetHandle($vueScan, "", "[CLASS:ComboBox; INSTANCE:46]")

	$statusBar 					      = ControlGetHandle($vueScan, "", "[CLASS:msctls_statusbar32; INSTANCE:1]")

	$tabPreviewScan 			    = ControlGetHandle($vueScan, "", "[CLASS:_wx_SysTabCtl32; INSTANCE:2]")
	$tabOptions 				      = ControlGetHandle($vueScan, "", "[CLASS:_wx_SysTabCtl32; INSTANCE:1]")

	$radioJpeg 					      = ControlGetHandle($vueScan, "", "[CLASS:Button; INSTANCE:24]")
	$radioPDF 					      = ControlGetHandle($vueScan, "", "[CLASS:Button; INSTANCE:30]")
	$radioTIFF 					      = ControlGetHandle($vueScan, "", "[CLASS:Button; INSTANCE:18]")

	$editJpegQuality 			    = ControlGetHandle($vueScan, "", "[CLASS:Edit; INSTANCE:84]")
	$editDefaultFolder 			  = ControlGetHandle($vueScan, "", "[CLASS:Edit; INSTANCE:96]")
	$editFileFormat 			    = ControlGetHandle($vueScan, "", "[CLASS:Edit; INSTANCE:98]")

	$editXSize 					      = ControlGetHandle($vueScan, "", "[CLASS:Edit; INSTANCE:58]")
	$editYSize 					      = ControlGetHandle($vueScan, "", "[CLASS:Edit; INSTANCE:59]")
	$editXOffset				      = ControlGetHandle($vueScan, "", "[CLASS:Edit; INSTANCE:60]")
	$editYOffset				      = ControlGetHandle($vueScan, "", "[CLASS:Edit; INSTANCE:61]")


EndFunc

Func setInstanceIDs()

	; Check the Options label
	$lblBitsPerPixel = ControlGetHandle($vueScan, "", "[CLASS:Static; INSTANCE:288]")

	If ControlGetText($vueScan, "", $lblBitsPerPixel) = "Bits per pixel: " Then Return

	MsgBox($MB_OK, "", "Scanner " & $currentScanner & ": Need to run setup... Press OK then wait for 'Setup complete' message")

	$comboOptions = ControlGetHandle($vueScan, "", "[CLASS:ComboBox; INSTANCE:1]")

	ControlCommand($vueScan, "", $comboOptions, "SelectString", 'Basic' )
	Sleep(4 * 1000)
	ControlCommand($vueScan, "", $comboOptions, "SelectString", 'Standard' )
	Sleep(4 * 1000)
	ControlCommand($vueScan, "", $comboOptions, "SelectString", 'Professional' )

	MsgBox($MB_OK, "", "Scanner " & $currentScanner & " setup complete!")

EndFunc

Func finish()
	wLog('Script finished (' & @exitCode & ')')
EndFunc

Func setupScanners()
	For $i = 0 to UBound($scanners) - 1
		setupScanner($i + 1)
	Next
EndFunc

Func setupScanner($scannerInstance)
	wLog('Setting up scanner ' & $scannerInstance)

	setCurrentScanner($scannerInstance)

	; Setup file output
	setupOutput()

	; Turn off the external viewer
	ControlCommand($vueScan, "", $comboExternalViewer, "SelectString", 'None' )

EndFunc

Func setupOutput()

	; Set crop size to manual
	ControlCommand($vueScan, "", $comboCropSize, "SelectString", 'Manual' )

	; Turn off PDF
	if ControlCommand($vueScan, "", $radioPDF, "IsChecked", "") = 1 Then
		ControlCommand($vueScan, "", $radioPDF, "UnCheck", "")
		Sleep(1 * 1000)
	EndIf

EndFunc

Func findScanners()
	Local $scannerList = WinList("[CLASS:wxWindowNR]")
	For $i = 1 To $scannerList[0][0]
		; Only want maximised scanners (best way to tell the difference)
		If $scannerList[$i][0] <> "" And BitAND(WinGetState($scannerList[$i][1]), 32) Then
			$scannerCount = $scannerCount + 1
			Redim $scanners[$scannerCount]
			$scanners[$i-1] = $scannerList[$i][1]
		EndIf
	Next

	if $scannerCount = 0 Then
		MsgBox($MB_SYSTEMMODAL, "", "No maximised scanners found")
		Exit
	ElseIf $scannerCount > 1 Then
		MsgBox($MB_SYSTEMMODAL, "", "More than 1 maximised scanner found")
		Exit
		; MsgBox($MB_SYSTEMMODAL, "", "Maximised scanners found: " & $scannerCount)
	EndIf
	_ArrayReverse($scanners)
	wLog('Finding maximised scanners. Scanners found: ' & $scannerCount)
EndFunc

Func setMode($mode)
  ControlCommand($vueScan, "", $comboMode, "SelectString", $mode)
  Sleep(5 * 1000)

	; Turn off infrared cleaning on transmission mode
	If $mode = $MODE_TRANSMISSION Then
		$comboInfraredClean	= ControlGetHandle($vueScan, "", "[CLASS:ComboBox; INSTANCE:46]")
		ControlCommand($vueScan, "", $comboInfraredClean, "SelectString", 'None' )
  EndIf

	; Set the preview resolution
	ControlCommand($vueScan, "", $comboPreviewResolutions, "SelectString", $PREVIEW_RESOLUTION & ' dpi')

	; Set the color balance
	ControlCommand($vueScan, "", $comboColorBalance, "SelectString", 'None')

	; Re-run the output setup
	setupOutput()

	; Get a preview
	preview()

	If $mode = $MODE_TRANSMISSION Then
	   $currentMode = 'T'
    Else
	   $currentMode = 'R'
    EndIf
EndFunc


Func preview()
	ControlClick($vueScan, "", $btnPreview)
	Sleep(2 * 1000)
	Do
		$currentStatus = StatusbarGetText($vueScan, "", 1)
		Sleep(2 * 1000)
	Until $currentStatus = 'Press Preview, adjust crop box, press Scan'
EndFunc

Func setOutputFolder($folder)
	wLog('Setting output folder to ' & $folder)
	$currentFolder = $folder
EndFunc

Func setSample($sample)
	wLog('Setting sample to ' & $sample)
	$currentSample = $sample
EndFunc

Func setOutputFilename($filename)
	ControlSetText($vueScan, "", $editFileFormat, "")
	ControlCommand($vueScan, "", $editFileFormat, "EditPaste", $filename)
EndFunc

Func selectRegion($xSize, $ySize, $xOffset, $yOffset)
	wLog('Selecting region ' & $xSize & ", " & $ySize & ", " & $xOffset & ", " & $yOffset)

	Local $currentTab = ControlCommand($vueScan, "", $tabPreviewScan, "CurrentTab")
	if $currentTab = 2 Then
		ControlCommand($vueScan, "", $tabPreviewScan, "TabLeft")
	EndIf
	Sleep(500)
	ControlSetText($vueScan, "", $editXSize, $xSize)
	ControlSetText($vueScan, "", $editYSize, $ySize)
	ControlSetText($vueScan, "", $editXOffset, $xOffset)
	ControlSetText($vueScan, "", $editYOffset, $yOffset)
    Sleep(1000)
EndFunc

Func setFileType($fileType, $imageType, $jpegQuality)

	If $fileType = $FILE_TYPE_JPEG Then
		; Turn off TIFF
		if ControlCommand($vueScan, "", $radioTIFF, "IsChecked", "") = 1 Then
			ControlCommand($vueScan, "", $radioTIFF, "UnCheck", "")
			Sleep(1 * 1000)
		EndIf

		; Turn on JPEG
		if ControlCommand($vueScan, "", $radioJpeg, "IsChecked", "") = 0 Then
			ControlCommand($vueScan, "", $radioJpeg, "Check", "")
			Sleep(1 * 1000)
		EndIf

		; Set the JPEG quality
		ControlSetText($vueScan, "", $editJpegQuality, $jpegQuality)

	Else

		; Turn off JPEG
		if ControlCommand($vueScan, "", $radioJpeg, "IsChecked", "") = 1 Then
			ControlCommand($vueScan, "", $radioJpeg, "UnCheck", "")
			Sleep(1 * 1000)
		EndIf

		; Turn on TIFF
		if ControlCommand($vueScan, "", $radioTIFF, "IsChecked", "") = 0 Then
			ControlCommand($vueScan, "", $radioTIFF, "Check", "")
			Sleep(1 * 1000)
		EndIf

		; Set TIFF settings
		ControlCommand($vueScan, "", $comboTiffFileType, "SelectString", $imageType )
		ControlCommand($vueScan, "", $comboTiffCompression, "SelectString", "On" )

	EndIf

EndFunc

Func scanAtResolution($resolution, $imageType, $fileType = $FILE_TYPE_JPEG, $filePrefix = "", $jpegQuality = 90)
	wLog('Scanning at resolution: ' & $resolution & ", image type:" & $imageType & ", jpeg qual.:" & $jpegQuality)

	setFileType($fileType, $imageType, $jpegQuality)

	; Output filename
	setOutputFilename($filePrefix & $OUTPUT_FILE_FORMAT)

	; Resolution and image type
	ControlCommand($vueScan, "", $comboResolutions, "SelectString", $resolution & ' dpi' )
	ControlCommand($vueScan, "", $comboImageTypes, "SelectString", $imageType )

	; Grey channel (if applicable)
	If $imageType = $IMAGE_TYPE_8_GREY Or $imageType = $IMAGE_TYPE_16_GREY Then
		ControlCommand($vueScan, "", $comboGreyChannel, "SelectString", $GREY_CHANNEL)
	EndIf

	; Create the output directory (if necessary)
	$outputDir = createOutputDir($resolution)
	ControlSetText($vueScan, "", $editDefaultFolder, "")
	ControlCommand($vueScan, "", $editDefaultFolder, "EditPaste", $outputDir)

	;~ if Not($currentImageType = $imageType) And Not($currentImageType = "") Then
	;~ 	Need to run another preview - sometimes vuescan locks the image type in grey / rgb, re-previewing refreshes this
	;~ 	preview()
	;~ EndIf

	; Scan
	ControlClick($vueScan, "", $btnScan)
	Sleep(3 * 1000)
	Local $hTimer = TimerInit()
	Do
		$currentStatus = StatusbarGetText($vueScan, "", 1)
		Sleep(2 * 1000)
		if TimerDiff($hTimer) > ($SCAN_MAX_TIME * 1000) Then
			fail()
		EndIf
	Until $currentStatus = 'Press Preview, adjust crop box, press Scan'
	wLog('Scan complete (' & TimerDiff($hTimer) & 'ms)')

	$currentImageType = $imageType
	$currentFileType = $fileType
EndFunc

Func attachToMaximisedScanner()
	findScanners()
	setOutputFolder($OUTPUT_FOLDER)
	setCurrentScanner(1)
	setupScanner(1)
EndFunc

Func createOutputDir($resolution)

	$outputDir = $currentFolder

	; Create output dir
	; e.g. C:\Data\Scans\Globulus
	if Not(FileExists($outputDir)) Then
		DirCreate($outputDir)
	EndIf

	; Sample
	; e.g. C:\Data\Scans\Globulus\Leaf1
	$outputDir = $outputDir & "\" & $currentSample

	if Not(FileExists($outputDir)) Then
		DirCreate($outputDir)
	EndIf

	; Mode
	; e.g. C:\Data\Scans\Globulus\Leaf1\Trans
	if $currentMode = 'T' Then
		$outputDir = $outputDir & "\Trans"
	Else
		$outputDir = $outputDir & "\Ref"
	EndIf

	if Not(FileExists($outputDir)) Then
		DirCreate($outputDir)
	EndIf

	; Resolution
	; e.g. C:\Data\Scans\Globulus\Leaf1\Trans\2400
	$outputDir = $outputDir & "\" & $resolution

	if Not(FileExists($outputDir)) Then
		DirCreate($outputDir)
	EndIf

	Return $outputDir

EndFunc

Func wLog($entry)
	_FileWriteLog($logFile, $entry)
EndFunc

Func fail()
	wLog('Scan failed')
	FileClose($logFile)
	Exit
EndFunc

Func die()
	$finished = True
	MsgBox($MB_SYSTEMMODAL, "", "AutoIT Script Finished")
	Exit 0
EndFunc
