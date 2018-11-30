Set ProgressUI = CreateObject("Microsoft.SMS.TsProgressUI")
ProgressUI.CloseProgressDialog
set sh = CreateObject("Wscript.Shell")
call sh.Run("select.hta", 1, True)