(vl-load-com)

;; thay bang RAW URL file PS o tren sau khi ban upload
(setq *psRaw* "https://raw.githubusercontent.com/YourUser/YourRepo/main/Install_Fonts_FromZip.ps1")

(defun _WriteDCL (path / f txt)
  (setq txt
"fonts_menu : dialog {
  label = \"Font Installer\";
  : boxed_column {
    label = \"ZIP URL (GitHub Releases)\";
    : edit_box { key = \"zipurl\"; width = 60; value = \"https://github.com/SurveyRoyal/Survey/releases/download/CAIDATFONT/FONTCAD.zip\"; }
  }
  : row {
    : button { key = \"all\";  label = \"Install ALL\"; is_default = true; }
    : button { key = \"acad\"; label = \"Install AUTOCAD\"; }
    : button { key = \"cancel\"; label = \"Cancel\"; is_cancel = true; }
  }
}"
  )
  (setq f (open path "w")) (write-line txt f) (close f)
)

(defun _RunPSAdmin (args / sh)
  (setq sh (vlax-create-object "Shell.Application"))
  (vlax-invoke sh 'ShellExecute "powershell.exe" args "" "runas" 1)
  (vlax-release-object sh)
)

(defun _AddSupportPath (path / cur upath)
  (setq cur (getenv "ACAD"))
  (if (not cur) (setq cur ""))
  (if (not (vl-string-search (strcase path) (strcase cur)))
    (progn (setq upath (strcat path ";" cur)) (setenv "ACAD" upath))
  )
)

(defun c:FONTDCL (/ dcl dclPath zipurl choice psCmd)
  (setq dclPath (strcat (getenv "TEMP") "\\fonts_menu.dcl"))
  (_WriteDCL dclPath)
  (setq dcl (load_dialog dclPath))
  (if (not (new_dialog "fonts_menu" dcl)) (progn (unload_dialog dcl) (exit)))

  (setq zipurl "https://github.com/SurveyRoyal/Survey/releases/download/CAIDATFONT/FONTCAD.zip")
  (set_tile "zipurl" zipurl)
  (action_tile "zipurl" "(setq zipurl $value)")
  (action_tile "all"   "(setq choice \"all\")(done_dialog 1)")
  (action_tile "acad"  "(setq choice \"acad\")(done_dialog 1)")
  (action_tile "cancel" "(done_dialog 0)")

  (if (= 1 (start_dialog))
    (progn
      (unload_dialog dcl)
      (if (or (not zipurl) (= zipurl "")) (progn (princ "\nNo ZIP URL.") (princ) (exit)))
      (cond
        ((= choice "all")
          (setq psCmd
            (strcat
              "-NoProfile -ExecutionPolicy Bypass -Command "
              "\"iwr -useb '" *psRaw* "' | iex; "
              "Install-Fonts_FromZip -ZipUrl '" zipurl "' -DoShx -DoTtf -DoPlot -OnlyNew -DestShx 'C:\\FONTCAD\\SHX'\""
            )
          )
          (_RunPSAdmin psCmd)
          (_AddSupportPath "C:\\FONTCAD\\SHX")
          (princ "\n> ALL: SHX->C:\\FONTCAD\\SHX, TTF->Windows, CTB/STB->Plot Styles. Support Path updated.")
        )
        ((= choice "acad")
          (setq psCmd
            (strcat
              "-NoProfile -ExecutionPolicy Bypass -Command "
              "\"iwr -useb '" *psRaw* "' | iex; "
              "Install-Fonts_FromZip -ZipUrl '" zipurl "' -DoShx -OnlyNew -DestShx 'C:\\FONTCAD\\SHX'\""
            )
          )
          (_RunPSAdmin psCmd)
          (_AddSupportPath "C:\\FONTCAD\\SHX")
          (princ "\n> AUTOCAD: Only SHX installed to C:\\FONTCAD\\SHX. Support Path updated.")
        )
      )
    )
    (unload_dialog dcl)
  )
  (princ)
)
