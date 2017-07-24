#cs ----------------------------------------------------------------------------
 AutoIt Version: 3.3.14.2
 Author:         OpenSourceOV.org (http://opensourceov.org)
#ce ----------------------------------------------------------------------------

; Procedure:
; 1. Open a Vuescan window for each scanner
; 2. Reset the scanner settings in each window by going to File > Default settings (must be done *before* setting the scanner)
; 3. Set the relevant scanner for each Vuescan window (Set window 1 = scanner 1, window 2 = scanner 2 etc)
; 4. Minimise all Vuescan windows
; 5. Maximise the Vuescan window you want to link with this script, then run this script.
; Note: the maximised Vuescan window doesn't have to be in the foreground i.e. you can maximise vuescan then switch to a folder to run the script.

; Config
; =============================================================
Const $SCAN_INTERVAL_SEC = 600
Const $END_DATE_TIME = "2016/09/09 17:00:00" ; Important - when the scanning should stop. Format YYYY/mm/dd HH:mm:ss
Const $LOG_FILE = "C:\Data\scanner_1_log.log"
Const $SCAN_MAX_TIME = 1200 ; maximum scan time before scan fail triggered, in seconds.
Const $OUTPUT_FOLDER = "C:\Data\05092016_glob_j"
Const $MODE_TRANSMISSION = 'Transparency'
Const $MODE_REFLECTIVE = 'Flatbed'
; =============================================================

#include "include.au3"

; Commands to run before the scanning procedure begins - useful for settings that are the same for all scans e.g. sample, mode etc
; =============================================================
Func preScan()
	attachToMaximisedScanner()

	; Pre-scan instructions defined here

EndFunc

; Commands to run once the scanning procedure has finished (leave empty in most cases)
; =============================================================
Func postScan()
EndFunc

; The scanning procedure, repeated with $SCAN_INTERVAL_SEC delay between cycles until $END_DATE_TIME
; =============================================================
Func scanningProcedure()

	; Example:

	;~ setSample('Leaf_1')
	;~ selectRegion(32.147,50.123,0,52.187)
	;~ scanAtResolution('2400', $IMAGE_TYPE_8_GREY, $FILE_TYPE_TIFF)

	;~ setSample('Leaf_2')
	;~ selectRegion(31.803,38.021,0.344,196.655)
	;~ scanAtResolution('2400', $IMAGE_TYPE_8_GREY, $FILE_TYPE_TIFF)

EndFunc

; Scan just a specified number of times, n, use scanRepeat(n) e.g.
; scanRepeat(2)

; Scan once then exit (useful for quick test) use scanOnce() e.g.
; scanOnce()

; Scan until $END_DATE_TIME
startScan()