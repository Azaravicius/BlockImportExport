;;-----------=={ Import or Export Blocks with attributs }==-------------;;
;;                                                                      ;;
;;  This program allows a user to import or export blocks in/from       ;;
;;  active drawing.                                                     ;;
;;                                                                      ;;
;;  Upon calling the program with 'BlockImportExport' at the command    ;;
;;  line, a new window will appear. Use it to import blocks to current  ;;
;;  drawing or export them from current drawing.                        ;;
;;                                                                      ;;
;;----------------------------------------------------------------------;;
;;  Author:  Donatas Azaravičius                                        ;;
;;----------------------------------------------------------------------;;
;;  Version 2.0    -    2022-02-20                                      ;;
;;----------------------------------------------------------------------;;
;;  Licence: European Union Public Licence (EUPL)                       ;;
;;  https://joinup.ec.europa.eu/collection/eupl/                        ;;
;;----------------------------------------------------------------------;;

;;----------------------------------------------------------------------;;
;;                        COMMAND FUNCTION                              ;;
;;----------------------------------------------------------------------;;
(defun C:BlockImportExport ()
  (vl-load-com)
  (command "_OPENDCL")
  (if (dcl-Project-Load "BlockImportExport" T "BlockIE")
    (dcl-Form-Show BlockIE/Main)
  )
  (princ)
)

;;----------------------------------------------------------------------;;
;;                    START C:BlockIE/Main CODE                         ;;
;;----------------------------------------------------------------------;;

(defun C:BlockIE/Main#OnInitialize (/ Delimiter)
  (dcl-Control-SetCaption BlockIE/Main/LabelCopyright "Created by Donatas Azaravičius. Licence EUPL. Software version: 2.0")
  
  ; Columns names for blocks data grid.
  (dcl-Grid-AddColumns BlockIE/Main/GridData 
    (list
      (list "Name" 0 100)
      (list "Y" 1 80)
      (list "X" 1 80)
      (list "Elevation" 1 70)
    )
  )
  
  ; Columns names for blocks summary grid.
  (dcl-Grid-AddColumns BlockIE/Main/GridInfo 
    (list
      (list "Name" 0 70)
      (list "Count" 1 50)
      (list "Comment" 0 200)
    )
  )
  
  ; Check if LastDirSelected exist, if not set it to C:\\
  (if (= (DA:GetCfgValue "LastDirSelected") "")
    (DA:SetCfgValue "LastDirSelected" "C:\\")
  )
  ; Check if ExternalDWGDir exist, if not set it to C:\\
  (if (= (DA:GetCfgValue "ExternalDWGDir") "")
    (DA:SetCfgValue "ExternalDWGDir" "C:\\")
  )
  ; Check if ExtensionCSV exist, if not set it to "1"
  (if (= (DA:GetCfgValue "ExtensionCSV") "")
    (DA:SetCfgValue "ExtensionCSV" "1")
  )
  ; Check if ExtensionTXT exist, if not set it to "0"
  (if (= (DA:GetCfgValue "ExtensionTXT") "")
    (DA:SetCfgValue "ExtensionTXT" "0")
  )
  ; Set delimiter
  (if (= (setq Delimiter (DA:GetCfgValue "Delimiter")) "")
    (progn
      (DA:SetCfgValue "Delimiter" "Auto")
      (DA:SetCfgValue "CustomDelimiter" "")
      (dcl-OptionList-SetCurSel BlockIE/Main/OptionDelimiter 0)
    )
    (cond
      ( (= Delimiter "Auto") (dcl-OptionList-SetCurSel BlockIE/Main/OptionDelimiter 0) )
      ( (= Delimiter "Semi-colon") (dcl-OptionList-SetCurSel BlockIE/Main/OptionDelimiter 1) )
      ( (= Delimiter "Comma") (dcl-OptionList-SetCurSel BlockIE/Main/OptionDelimiter 2) )
      ( (= Delimiter "Space") (dcl-OptionList-SetCurSel BlockIE/Main/OptionDelimiter 3) )
      ( (= Delimiter "Tab") (dcl-OptionList-SetCurSel BlockIE/Main/OptionDelimiter 4) )
      ( (= Delimiter "Custom") 
        (dcl-OptionList-SetCurSel BlockIE/Main/OptionDelimiter 5) 
        (dcl-Control-SetEnabled BlockIE/Main/TextBoxDelimiter T)
        (dcl-Control-SetText BlockIE/Main/TextBoxDelimiter (DA:GetCfgValue "CustomDelimiter"))
      )
    )
  )
)

; Double clicking on block in block list of current drawing opens block preview window
(defun C:BlockIE/Main/BlockListCurrent#OnDblClicked (Item /)
  (dcl-Form-Show BlockIE/BlockPreview)
)

; Pressing "Import External Blocks" button opens window 
; for importing new blocks to current drawing from external drawing.
(defun c:BlockIE/Main/BtnExternalBlocks#OnClicked ()
  (dcl-Form-Show BlockIE/External)
)

; Enables or disables custom delimiter text box.
(defun C:BlockIE/Main/OptionDelimiter#OnSelChanged (ItemIndexOrCount Value /)
  (if (= Value "Custom")
    (progn
      (dcl-Control-SetEnabled BlockIE/Main/TextBoxDelimiter T)
      (dcl-Control-SetText BlockIE/Main/TextBoxDelimiter (DA:GetCfgValue "CustomDelimiter"))
    )
    (dcl-Control-SetEnabled BlockIE/Main/TextBoxDelimiter NIL)
  )
  (DA:SetCfgValue "Delimiter" Value)
)

; Opens new window for importing block data from text file.
; Also disables Export button, so user would know he is
; importing block data to drawing and enables "Ok" button.
(defun C:BlockIE/Main/BtnImport#OnClicked ()
  (dcl-Form-Show BlockIE/Import)
)

; Opens new window for importing block data from clipboard.
; Also disables Export button, so user would know he is
; importing block data to drawing and enables "Ok" button.
(defun C:BlockIE/Main/BtnPast#OnClicked ()
  (dcl-Form-Show BlockIE/Past)
)

; Opens new window for exporting block data from current drawing.
(defun C:BlockIE/Main/BtnExport#OnClicked ()
  (dcl-Form-Show BlockIE/Export)
)

; Opens new window with formated blocks data from blocks grid
; for fast copying to excel.
(defun C:BlockIE/Main/BtnCopy#OnClicked ()
  (dcl-Form-Show BlockIE/Copy)
)

; If Export button is Enabled export block data from block grid to csv file.
; Otherwise import blocks to drawing from block grid data.
(defun C:BlockIE/Main/BtnOk#OnClicked (/ RowCount Index BlockName BlockExistList RowData)
  (if (dcl-Control-GetEnabled BlockIE/Main/BtnExport)
    (DA:SaveCSV (DA:MakeFormatedString))
    (progn
      (setq
        RowCount (dcl-Grid-GetRowCount BlockIE/Main/GridData)
        Index 0
        BlockExistList (list)
      )
      (while (> RowCount Index)
        (setq 
          RowData (dcl-Grid-GetRowCells BlockIE/Main/GridData Index)
          BlockName (nth 0 RowData)
        )
        (if (not (assoc BlockName BlockExistList))
          (setq BlockExistList (cons (cons BlockName (DA:BlockExist BlockName)) BlockExistList))
        )
        (if (cdr (assoc BlockName BlockExistList))
          (DA:InsertBlock BlockName (list (atof (nth 1 RowData)) (atof (nth 2 RowData)) (atof (nth 3 RowData))) (DA:SubList RowData 4 nil))
        )
        (setq Index (1+ Index))
      )
    )
  )
  (dcl-Form-Close BlockIE/Main 3)
)

; Just close Block import-Export window.
(defun C:BlockIE/Main/BtnCancel#OnClicked ()
  (dcl-Form-Close BlockIE/Main 2)
)

;;----------------------------------------------------------------------;;
;;               START C:BlockIE/BlockPreview CODE                      ;;
;;----------------------------------------------------------------------;;

(defun C:BlockIE/BlockPreview#OnInitialize ()
  (dcl-Control-SetBlockName BlockIE/BlockPreview/BlockView (dcl-BlockList-GetItemText BlockIE/Main/BlockListCurrent (dcl-BlockList-GetCurSel BlockIE/Main/BlockListCurrent)))
  (dcl-BlockView-Zoom BlockIE/BlockPreview/BlockView 0.5)
)

;;----------------------------------------------------------------------;;
;;                    START C:BlockIE/Past CODE                         ;;
;;----------------------------------------------------------------------;;

(defun C:BlockIE/Past#OnInitialize ()
  (dcl-Control-SetText BlockIE/Past/TB_Past "")
)

; Import data from Past window to block data grid
(defun C:BlockIE/Past/BtnOK#OnClicked (/ Data Delimiter FirstLine Line)
  (setq 
    Data (dcl-Control-GetText BlockIE/Past/TB_Past)
    Delimiter (dcl-OptionList-GetButtonCaption BlockIE/Main/OptionDelimiter (dcl-OptionList-GetCurSel BlockIE/Main/OptionDelimiter))
  )
  (if (vl-string-search "\r\n" Data)
    (setq Data (DA:StrToLst Data "\r\n"))
    (setq Data (DA:StrToLst Data "\n"))
  )
  (setq FirstLine (car Data))
  (if (= "Auto" Delimiter)
    (cond
      ((vl-string-search "\t" FirstLine) (setq Delimiter "\t"))
      ((vl-string-search ";" FirstLine) (setq Delimiter ";"))
      ((vl-string-search "," FirstLine) (setq Delimiter ","))
      ((vl-string-search " " FirstLine) (setq Delimiter " "))
      (T (setq Delimiter NIL))
    )
    (cond
      ((= "Semi-colon" Delimiter) (setq Delimiter ";"))
      ((= "Comma" Delimiter) (setq Delimiter ","))
      ((= "Space" Delimiter) (setq Delimiter " "))
      ((= "Tab" Delimiter) (setq Delimiter "\t"))
      ((= "Custom" Delimiter) (setq Delimiter (dcl-Control-GetText BlockIE/Main/TextBoxDelimiter)))
      (T (setq Delimiter NIL))
    )
  )
  (if Delimiter 
    (progn
      (DA:UpdateGridAttr (- (DA:CountLetter FirstLine Delimiter) 3))
      (foreach Line Data
        (if (/= Line "") (dcl-Grid-AddString BlockIE/Main/GridData Line Delimiter))
      )
    )
    (dcl-MessageBox "Could not get delimiter!" "ERROR" 2 4)
  )
  (dcl-Form-Close BlockIE/Past 3)
  (dcl-Control-SetEnabled BlockIE/Main/BtnExport NIL)
  (dcl-Control-SetEnabled BlockIE/Main/BtnOk T)
  (DA:ProcessBlocks)
)

(defun C:BlockIE/Past/BtnCancel#OnClicked (/)
  (dcl-Form-Close BlockIE/Past 2)
)

;;----------------------------------------------------------------------;;
;;                   START C:BlockIE/Copy CODE                         ;;
;;----------------------------------------------------------------------;;

; Formates blocks grid data for copying
(defun C:BlockIE/Copy#OnInitialize ()
  (dcl-Control-SetText BlockIE/Copy/TB_Copy (DA:MakeFormatedString))
)

;;----------------------------------------------------------------------;;
;;                  START C:BlockIE/Import CODE                        ;;
;;----------------------------------------------------------------------;;

(defun C:BlockIE/Import#OnInitialize (/)
  (dcl-Control-SetValue BlockIE/Import/CB_CSV (atoi (DA:GetCfgValue "ExtensionCSV")))
  (dcl-Control-SetValue BlockIE/Import/CB_TXT (atoi (DA:GetCfgValue "ExtensionTXT")))
  (DA:ListDirContent (DA:GetCfgValue "LastDirSelected"))
)

; Lets user select directory with blocks data files.
; If new directory selected needs to update paths of currently selected 
; files with blocks data.
(defun C:BlockIE/Import/BtnBrowse#OnClicked (/ OldPath NewPath)
  (setq 
    NewPath (DA:BrowseForFolder "Select Folder" "" 0)
    OldPath (dcl-Control-GetText BlockIE/Import/TB_Directory)
  )
  (if NewPath
    (progn
      (DA:SetCfgValue "LastDirSelected" NewPath)
      (if (> (dcl-ListBox-GetCount BlockIE/Import/LB_Selected) 0)
        (DA:UpdateFilePaths OldPath NewPath)
      )
      (DA:ListDirContent NewPath)
    )
  )
)

; Controls action of double clicking on directory view item.
; If item with index 0 is double clicked it means to go one folder up.
; To do this we need to updated paths of currently selected files.
; If item index is not 0, but item text have \ simbol it means directory was clicked.
; We need to enter this directory and update paths of currently selected files.
; If items index is not 0 and item text don't contane \ simbol, it means file was clicked.
; We need to remove selected file from directory view and add it to currently selected file list.
(defun C:BlockIE/Import/LB_DirView#OnDblClicked (/ Index SelectedFiles Position LastDir ItemText File LastDirSelected)
  (setq 
    Index (dcl-ListBox-GetCurSel BlockIE/Import/LB_DirView)
    SelectedFiles (dcl-Control-GetList BlockIE/Import/LB_Selected)
    LastDirSelected (DA:GetCfgValue "LastDirSelected")
  )
  (if (= Index 0)
    (progn
      (setq 
        Position (vl-string-position (ascii "\\") LastDirSelected nil T)
        LastDir (substr LastDirSelected (+ Position 2))
        LastDirSelected (substr LastDirSelected 1 Position)
      )
      (dcl-ListBox-Clear BlockIE/Import/LB_Selected)
      (foreach File SelectedFiles
        (if (vl-string-search "..\\" File)
          (dcl-ListBox-AddString BlockIE/Import/LB_Selected (substr File 4))
          (dcl-ListBox-AddString BlockIE/Import/LB_Selected (strcat LastDir "\\" File))
        )
      )
      (DA:ListDirContent LastDirSelected)
    )
    (progn
      (setq 
        ItemText (dcl-ListBox-GetItemText BlockIE/Import/LB_DirView Index)
        Position (vl-string-position (ascii "\\") ItemText)
      )
      (if Position
        (progn
          (setq 
            ItemText (substr ItemText 1 Position)
            LastDirSelected (strcat LastDirSelected "\\" ItemText)
          )
          (dcl-ListBox-Clear BlockIE/Import/LB_Selected)
          (foreach File SelectedFiles
            (if (= (vl-string-search ItemText File) 0)
              (dcl-ListBox-AddString BlockIE/Import/LB_Selected (substr File (+ 2 (strlen ItemText))))
              (dcl-ListBox-AddString BlockIE/Import/LB_Selected (strcat "..\\" File))
            )
          )
          (DA:ListDirContent LastDirSelected)
        )
        (progn
          (dcl-ListBox-DeleteItem BlockIE/Import/LB_DirView Index)
          (dcl-ListBox-AddString BlockIE/Import/LB_Selected ItemText)
        )
      )
    )
  )
  (DA:SetCfgValue "LastDirSelected" LastDirSelected)
)

; Remove selected file from currently selected files list.
(defun C:BlockIE/Import/LB_Selected#OnDblClicked ()
  (DA:RemoveSelectedFiles (dcl-ListBox-GetItemText BlockIE/Import/LB_Selected (dcl-ListBox-GetCurSel BlockIE/Import/LB_Selected)))
)

(defun C:BlockIE/Import/BtnCancel#OnClicked ()
  (dcl-Form-Close BlockIE/Import 2)
)

; Move all selected files from directory view to selected files list.
(defun C:BlockIE/Import/BtnAdd#OnClicked (/ SelectedItems Item)
  (if (setq SelectedItems (dcl-ListBox-GetSelectedItems BlockIE/Import/LB_DirView))
    (foreach Item SelectedItems
      (DA:MoveSelectedFiles Item)
    )
  )
)

; Remove selected files from selected files list.
(defun C:BlockIE/Import/BtnRemove#OnClicked (/ SelectedFiles File)
  (if (setq SelectedFiles (dcl-ListBox-GetSelectedItems BlockIE/Import/LB_Selected))
    (foreach File SelectedFiles
      (DA:RemoveSelectedFiles File)
    )
  )
)

; Re-new directory view by CSV filter.
(defun C:BlockIE/Import/CB_CSV#OnClicked (Value /)
  (DA:SetCfgValue "ExtensionCSV" (itoa Value))
  (DA:ListDirContent (DA:GetCfgValue "LastDirSelected"))
)

; Re-new directory view by TXT filter.
(defun C:BlockIE/Import/CB_TXT#OnClicked (Value /)
  (DA:SetCfgValue "ExtensionTXT" (itoa Value))
  (DA:ListDirContent (DA:GetCfgValue "LastDirSelected"))
)

; Gets data from all selected files and imports data to block grid.
(defun C:BlockIE/Import/BtnOk#OnClicked (/ File Line FileData)
  (setq FileData (list) )
  (foreach File (dcl-Control-GetList BlockIE/Import/LB_Selected)
    (if (setq File (open (DA:MakeFullPath (dcl-Control-GetText BlockIE/Import/TB_Directory) File) "r"))
      (progn
        (while (setq Line (read-line File))
          (setq FileData (cons Line FileData))
        )
        (close File)
        (DA:AddDataToGrid FileData)
        (setq FileData (list))
      )
    )
  )
  (DA:ProcessBlocks)
  (dcl-Form-Close BlockIE/Import 3)
  (dcl-Control-SetEnabled BlockIE/Main/BtnExport NIL)
  (dcl-Control-SetEnabled BlockIE/Main/BtnOk T)
)

; Change directory view if user directly modify path in TB_Directory.
(defun C:BlockIE/Import/TB_Directory#OnReturnPressed (/ OldPath NewPath)
  (setq 
    NewPath (dcl-Control-GetText BlockIE/Import/TB_Directory)
    OldPath (DA:GetCfgValue "LastDirSelected")
  )
  (if NewPath
    (progn
      (if (= (substr NewPath (strlen NewPath)) "\\")
        (setq NewPath (substr NewPath 1 (1- (strlen NewPath))))
      )
      (if (> (dcl-ListBox-GetCount BlockIE/Import/LB_Selected) 0)
        (DA:UpdateFilePaths OldPath NewPath)
      )
      (DA:ListDirContent NewPath)
      (DA:SetCfgValue "LastDirSelected" NewPath)
    )
    (dcl-Control-SetText BlockIE/Import/TB_Directory OldPath)
  )
)

; Prevent BlockIE/Import form to close if user pressed "Enter"
(defun C:BlockIE/Import#OnCancelClose (Reason /)
  (if (= Reason 0) (setq Reason T))
)

;;----------------------------------------------------------------------;;
;;                  START C:BlockIE/Export CODE                        ;;
;;----------------------------------------------------------------------;;

; Select all blocks in drawing by block name for exporting and 
; populate block data grid using block data.
(defun C:BlockIE/Export/BtnByName#OnClicked (/ ssBlocks Index BlockData BlockDataString Item BlockDataCount Entity)
  (dcl-Form-Close BlockIE/Export 3)
  (dcl-Form-Hide BlockIE/Main T)
  (setq Index 0)
  (cond
    ( (not (setq Entity (entsel)))
      (princ "\nEntity not selected. Selection aborted.")
    )
    ( (setq ssBlocks (ssget "_X" (list (cons 2 (cdr (assoc 2 (entget (car Entity))))) (cons 410 "Model"))))
      (repeat (sslength ssBlocks)
        (setq 
          BlockData (DA:GetBlockData (ssname ssBlocks Index))
          BlockDataCount (length BlockData)
          Index (1+ Index)
          BlockDataString (strcat (nth 0 BlockData) "\t" (rtos (nth 1 BlockData)) "\t" (rtos (nth 2 BlockData)) "\t" (rtos (nth 3 BlockData)))
          BlockData (DA:SubList BlockData 4 NIL)
        )
        (foreach Item BlockData
          (setq BlockDataString (strcat BlockDataString "\t" Item))
        )
        (DA:UpdateGridAttr (- BlockDataCount 4))
        (dcl-Grid-AddString BlockIE/Main/GridData BlockDataString "\t")
      )
      (dcl-Control-SetEnabled BlockIE/Main/BtnImport NIL)
      (dcl-Control-SetEnabled BlockIE/Main/BtnPast NIL)
      (dcl-Control-SetEnabled BlockIE/Main/BtnOk T)
      (DA:ProcessBlocks)
    )
  )
  (dcl-Form-Hide BlockIE/Main NIL)
)

; Lets user select blocks for export and
; populate block data grid using block data.
(defun C:BlockIE/Export/BtnArea#OnClicked (/ ssBlocks Index BlockData BlockDataString Item BlockDataCount)
  (dcl-Form-Close BlockIE/Export 3)
  (dcl-Form-Hide BlockIE/Main T)
  (setq Index 0)
  (if (setq ssBlocks (ssget '((0 . "INSERT") (410 . "Model"))))
    (progn
      (repeat (sslength ssBlocks)
        (setq 
          BlockData (DA:GetBlockData (ssname ssBlocks Index))
          BlockDataCount (length BlockData)
          Index (1+ Index)
          BlockDataString (strcat (nth 0 BlockData) "\t" (rtos (nth 1 BlockData)) "\t" (rtos (nth 2 BlockData)) "\t" (rtos (nth 3 BlockData)))
          BlockData (DA:SubList BlockData 4 NIL)
        )
        (foreach Item BlockData
          (setq BlockDataString (strcat BlockDataString "\t" Item))
        )
        (DA:UpdateGridAttr (- BlockDataCount 4))
        (dcl-Grid-AddString BlockIE/Main/GridData BlockDataString "\t")
      )
      (dcl-Control-SetEnabled BlockIE/Main/BtnImport NIL)
      (dcl-Control-SetEnabled BlockIE/Main/BtnPast NIL)
      (dcl-Control-SetEnabled BlockIE/Main/BtnOk T)
      (DA:ProcessBlocks)
    )
  )
  (dcl-Form-Hide BlockIE/Main NIL)
)


;;----------------------------------------------------------------------;;
;;                 START C:BlockIE/External CODE                       ;;
;;----------------------------------------------------------------------;;

(defun C:BlockIE/External#OnInitialize (/)
  (dcl-Control-SetText BlockIE/External/TB_File "Current drawing")
)

; Opens window for selecting drawing file with external blocks.
(defun C:BlockIE/External/BtnBrowse#OnClicked (/ dwg)
  (cond
    ( (not (setq dwg (getfiled "Select Source Drawing" (DA:GetCfgValue "ExternalDWGDir") "dwg;dwt;dws" 16)))
      (princ "\nCancel")
    )
    ( T
      (dcl-Control-SetText BlockIE/External/TB_File dwg)
      (dcl-BlockList-LoadDwg BlockIE/External/BlockList dwg)
      (dcl-Control-SetEnabled BlockIE/External/BtnOk T)
      (DA:SetCfgValue "ExternalDWGDir" (strcat (vl-filename-directory dwg) "\\"))
    )
  )
)

(defun C:BlockIE/External/BtnCancel#OnClicked (/)
  (dcl-Form-Close BlockIE/External 2)
)

; Inserts external dwg file to current drawing file
; to get all blocks from it to current drawing and delete it.
(defun C:BlockIE/External/BtnOk#OnClicked ()
  (vla-Delete 
    (vla-InsertBlock 
      (vla-get-ModelSpace 
        (vla-get-activedocument 
          (vlax-get-acad-object)
        )
      ) 
      (vlax-3D-point 0 0 0)	
      (dcl-Control-GetText BlockIE/External/TB_File) 1 1 1 0
    )
  )
  (dcl-Form-Close BlockIE/External 3)
  (DA:ProcessBlocks)
)

;;----------------------------------------------------------------------;;
;;                   START EXTRA FUNCTIONS CODE                         ;;
;;----------------------------------------------------------------------;;

; Split a string using a given delimiter into list.
; String    - [str] String to Split.
; Delimiter - [str] Delimiter by which to split string.
; Returns:    [list] List of strings.
(defun DA:StrToLst (String Delimiter / Lenght Data Position)
  (setq 
    Lenght (1+ (strlen Delimiter))
    Data (list)
  )
  (while (setq Position (vl-string-search Delimiter String))
    (setq
      Data (cons (substr String 1 Position) Data)
      String (substr String (+ Position Lenght))
    )
  )
  (reverse (cons String Data))
)

; Count appearance of Letter in a String.
; String - [str] String to be searched.
; Letter - [str] Letter to count.
; Returns: [long] Times Letter appeared in String.
(defun DA:CountLetter (String Letter / Data)
  (setq Data (vl-string->list String))
  (- (length Data) (length (vl-remove (ascii Letter) Data)))
)

; Displays a dialog prompting the user to select a folder.
; Message   - [str] message to display at top of dialog
; Directory - [str] root directory (or "")
; Flag      - [int] bit-coded flag specifying dialog display settings
; Returns: [str] Selected folder path, else nil.
(defun DA:BrowseForFolder (Message Directory Flag / Shell HWND Browser Self Path)
  (setq
    Shell (vla-getinterfaceobject (vlax-get-acad-object) "shell.application")
    HWND (vl-catch-all-apply 'vla-get-hwnd (list (vlax-get-acad-object)))
    Browser (vlax-invoke-method Shell 'browseforfolder HWND Message Flag Directory)
  )
  (if Browser
    (setq
      Self (vlax-get-property Browser 'self)
      Path (vl-string-right-trim "\\" (vl-string-translate "/" "\\" (vlax-get-property Self 'path)))
    )
  )
  (if Self (vlax-release-object Self))
  (if Browser (vlax-release-object Browser))
  (if Shell (vlax-release-object Shell))
  Path
)

; substr function analog for list
; Data   - [list] List from which sublist is to be returned
; Start  - [int] Zero-based index at which to start the sublist
; Lenght - [int] Length of the sublist or nil to return all items following idx
; Returns: [list] Sublist of list.
(defun DA:SubList (Data Start Lenght / Return)
  (setq 
    Lenght  (if Lenght 
              (min Lenght (- (length Data) Start)) 
              (- (length Data) Start)
            )
    Start (+ Start Lenght)
  )
  (repeat Lenght
    (setq Return (cons (nth (setq Start (1- Start)) Data) Return))
  )
)

; Function to process all data in Blocks data grid. 
; It needs to be called after data is added to Block data grid.
; This function will count how many times same block exist in Block data grid
; and if block exist in current drawing. 
; All this inforamtion will be added to Block information grid.
(defun DA:ProcessBlocks (/ ColumnData CellData DottedList DottedPair RowData)
  (setq 
    ColumnData (dcl-Grid-GetColumnCells BlockIE/Main/GridData 0)
    DottedList (list)
  )
  (foreach CellData ColumnData
    (setq DottedPair (assoc CellData DottedList))
    (if (= NIL DottedPair)
      (setq DottedList (cons (cons CellData 1) DottedList ))
      (setq DottedList 
        (subst 
          (cons (car DottedPair) (1+ (cdr DottedPair)))
          DottedPair
          DottedList
        )
      )
    )
  )
  (dcl-Grid-Clear BlockIE/Main/GridInfo)
  (foreach RowData DottedList
    (dcl-Grid-AddRow BlockIE/Main/GridInfo 
      (car RowData) 
      (itoa (cdr RowData)) 
      (if (= NIL (tblsearch "BLOCK" (car RowData)))
        "Block was not found in current drawing!"
        ""
      )
    )
  )
)

; In "Block data import" window adds folders and files to directory view list.
; Path - [str] directory path to display content.
(defun DA:ListDirContent(Path / Dirs Files Item ListOfSelectedFiles FileSelected NotInList ExtensionCSV ExtensionTXT Extension)
  (setq 
    Dirs (vl-directory-files Path nil -1)
    Files (vl-directory-files Path nil 1)
    ListOfSelectedFiles (dcl-Control-GetList BlockIE/Import/LB_Selected)
    NotInList T
    ExtensionCSV (dcl-Control-GetValue BlockIE/Import/CB_CSV)
    ExtensionTXT (dcl-Control-GetValue BlockIE/Import/CB_TXT)
  )
  (dcl-Control-SetText BlockIE/Import/TB_Directory Path)
  (dcl-ListBox-Clear BlockIE/Import/LB_DirView)
  (dcl-ListBox-AddString BlockIE/Import/LB_DirView "..")
  (foreach Item Dirs
    (if (and (/= Item ".") (/= Item ".."))
      (dcl-ListBox-AddString BlockIE/Import/LB_DirView (strcat Item "\\"))
    )
  )
  (foreach Item Files
    (if ListOfSelectedFiles
      (foreach FileSelected ListOfSelectedFiles
        (if (= (vl-string-search Item FileSelected) 0)
          (setq NotInList nil)
        )
      )
    )
    (if NotInList 
      (progn
        (setq Extension (vl-filename-extension Item))
        (cond
          ((and (= ExtensionCSV 1) (= Extension ".csv")) (dcl-ListBox-AddString BlockIE/Import/LB_DirView Item))
          ((and (= ExtensionTXT 1) (= Extension ".txt")) (dcl-ListBox-AddString BlockIE/Import/LB_DirView Item))
          ((and (/= ExtensionCSV 1) (/= ExtensionTXT 1)) (dcl-ListBox-AddString BlockIE/Import/LB_DirView Item))
        )
      )
    )
    (setq NotInList T)
  )
)

; Updates selected files path after changeing current directory.
; OldPath - [str] Path to previous directory.
; NewPath - [str] Path to new current directory.
(defun DA:UpdateFilePaths (OldPath NewPath / ListOfSelectedFiles FileSelected FilePathLength PathCopy PathSegment FilePath NewListOfSelectedFiles)
  (setq 
    OldPath (DA:StrToLst OldPath "\\")
    NewPath (DA:StrToLst NewPath "\\")
    ListOfSelectedFiles (dcl-Control-GetList BlockIE/Import/LB_Selected)
    FilePath ""
    NewListOfSelectedFiles (list)
  )
  (foreach FileSelected ListOfSelectedFiles
    (setq 
      FileSelected (DA:StrToLst FileSelected "\\")
      FilePathLength (length FileSelected)
      PathCopy (reverse OldPath)
    )
    (while (> FilePathLength 0)
      (setq FilePathLength (1- FilePathLength))
      (if (= ".." (car FileSelected))
        (progn
          (setq 
            FileSelected (cdr FileSelected)
            PathCopy (cdr PathCopy)
          )
        )
        (progn
          (setq FilePathLength 0)
          (if (not (vl-string-position (ascii ":") (car FileSelected)))
            (setq FileSelected (append (reverse PathCopy) FileSelected))
          )
        )
      )
    )
    (if (= (car NewPath) (car FileSelected))
      (progn
        (setq 
          FilePathLength (length NewPath)
          PathCopy NewPath
        )
        (while (> FilePathLength 0)
          (setq FilePathLength (1- FilePathLength))
          (if (= (car PathCopy) (car FileSelected))
            (setq FileSelected (cdr FileSelected))
            (progn
              (foreach PathSegment PathCopy
                (setq FileSelected (cons ".." FileSelected))
              )
              (setq FilePathLength 0)
            )
          )
          (setq PathCopy (cdr PathCopy))
        )
      )
    )
    (setq FileSelected (reverse FileSelected))
    (foreach PathSegment FileSelected
      (setq FilePath (strcat PathSegment "\\" FilePath))
    )
    (setq 
      FilePath (substr FilePath 1 (1- (strlen FilePath)))
      NewListOfSelectedFiles (cons FilePath NewListOfSelectedFiles)
      FilePath ""
    )
  )
  (dcl-ListBox-Clear BlockIE/Import/LB_Selected)
  (dcl-ListBox-AddList BlockIE/Import/LB_Selected newListOfSelectedFiles)
)

; Moves selected file from directory view list to currently selected files list.
; Text - [str] Name of selected item.
; Selected item can't have id 0, because it is not file but move up directory.
; Selected item can't have "\" because only folders can have it.
(defun DA:MoveSelectedFiles (Text / Id)
  (setq Id (dcl-ListBox-FindStringExact BlockIE/Import/LB_DirView Text))
  (if (/= Id 0)
    (progn
      (if (not (vl-string-position (ascii "\\") text))
        (progn
          (dcl-ListBox-DeleteItem BlockIE/Import/LB_DirView Id)
          (dcl-ListBox-AddString BlockIE/Import/LB_Selected Text)
        )
      )
    )
  )
)

; Removes selected file from selected files list and if filename don't have "\"
; this file belongs to current directory, so it is returned to current directory.
; File - [str] Name of selected file from selected file list.
(defun DA:RemoveSelectedFiles (File /)
  (if (not (vl-string-search "\\" File))
      (dcl-ListBox-AddString BlockIE/Import/LB_DirView File)
  )
  (dcl-ListBox-DeleteItem BlockIE/Import/LB_Selected (dcl-ListBox-FindStringExact BlockIE/Import/LB_Selected File))
)

; Joins file Path and File to make full path to file.
; Path - [str] Path part to file.
; File - [str] File name with partial path.
; Returns: [str] Full path to file.
(defun DA:MakeFullPath (Path File / FullPath FilePathLength PathSegment)
  (setq 
    Path (reverse (DA:StrToLst Path "\\"))
    File (DA:StrToLst File "\\")
    FullPath ""
    FilePathLength (length File)
  )
  (while (> FilePathLength 0)
    (setq FilePathLength (1- FilePathLength))
    (if (= ".." (car File))
      (progn
        (setq 
          File (cdr File)
          Path (cdr Path)
        )
      )
      (progn
        (setq FilePathLength 0)
        (if (not (vl-string-position (ascii ":") (car File)))
          (setq File (append (reverse Path) File))
        )
      )
    )
  )
  (setq File (reverse File))
  (foreach PathSegment File
    (setq FullPath (strcat PathSegment "\\" FullPath))
  )
  (substr FullPath 1 (1- (strlen FullPath)))
)

; Will add block data to Block data grid.
; Data - [list] List of lists with all block data.
(defun DA:AddDataToGrid (Data / FirstLine Delimiter DelimiterCount Line)
  (setq 
    FirstLine (car Data)
    Delimiter (dcl-OptionList-GetButtonCaption BlockIE/Main/OptionDelimiter (dcl-OptionList-GetCurSel BlockIE/Main/OptionDelimiter))
  )
  (if (= "Auto" Delimiter)
    (cond
      ((vl-string-search "\t" FirstLine) (setq Delimiter "\t"))
      ((vl-string-search ";" FirstLine) (setq Delimiter ";"))
      ((vl-string-search "," FirstLine) (setq Delimiter ","))
      ((vl-string-search " " FirstLine) (setq Delimiter " "))
      (T (setq Delimiter NIL))
    )
    (cond
      ((= "Semi-colon" Delimiter) (setq Delimiter ";"))
      ((= "Comma" Delimiter) (setq Delimiter ","))
      ((= "Space" Delimiter) (setq Delimiter " "))
      ((= "Tab" Delimiter) (setq Delimiter "\t"))
      ((= "Custom" Delimiter) 
        (setq Delimiter (dcl-Control-GetText BlockIE/Main/TextBoxDelimiter))
        (DA:SetCfgValue "CustomDelimiter" Delimiter)
      )
      (T (setq Delimiter NIL))
    )
  )
  (if Delimiter 
    (progn
      (DA:UpdateGridAttr (- (DA:CountLetter FirstLine Delimiter) 3))
      (foreach Line Data
        (if (/= Line "") (dcl-Grid-AddString BlockIE/Main/GridData Line Delimiter))
      )
    )
    (dcl-MessageBox "Could not get delimiter!" "ERROR" 2 4)
  )
)

; Inserts block to model space.
; BlockName   - [str] Name of block to insert.
; Coordinates - [list] List of block coordinates (Y X Z)
; Attributes  - [list] List of block attributs or nil.
(defun DA:InsertBlock (BlockName Coordinates Attributes / Mspace BlRefObj)
  (setq 
    Mspace (vla-get-modelspace (vla-get-activedocument (vlax-get-acad-object)))
    BlRefObj (vla-InsertBlock Mspace (vlax-3d-point Coordinates) BlockName 1 1 1 0)
  )
  (if Attributes
    (mapcar
      '(lambda (a v) (vla-put-textstring a v))
      (vlax-safearray->list (vlax-variant-value (vla-getattributes BlRefObj)))
      Attributes
    )
  )
)

; Creates formated string from data in Block data grid.
; Returns: [str] All data from Block data grid as formated string.
(defun DA:MakeFormatedString (/ RowCount Index RowData Data CellData FirstCollumn)
  (setq
    RowCount (dcl-Grid-GetRowCount BlockIE/Main/GridData)
    Index 0
    Data ""
    FirstCollumn T
  )
  (while (> RowCount Index)
    (setq RowData (dcl-Grid-GetRowCells BlockIE/Main/GridData Index))
    (foreach CellData RowData
      (if FirstCollumn
        (progn
          (setq 
            Data (strcat Data CellData)
            FirstCollumn nil
          )
        )
        (setq Data (strcat Data "\t" CellData))
      )
    )
    (setq 
      Data (strcat Data "\r\n")
      FirstCollumn T
      Index (1+ Index)
    )
  )
  Data
)

; Saves data to file.
; Data - [str] Data to save to csv file.
(defun DA:SaveCSV (Data / File)
  (setq File (getfiled "Save CSV to" (strcat (vl-filename-directory (DA:GetCfgValue "LastDirSelected")) "\\") "csv" 1))
  (if (setq File (open File "w"))
    (progn
      (write-line Data File)
      (close File)
    )
    (princ "\n ERROR - Could not open file for write.")
  )
)

; Check if block exist in BLOCK table.
; Name - [str] Block name to search.
; If block exist return T
; If block don't exist return NIL
(defun DA:BlockExist (Name /)
  (not (= (tblsearch "BLOCK" Name) nil))
)

; Gets block data of existing block in drawing.
; BlockObject - [<Entity name:>] Entity name of the block being queried.
; Returns: [list] List of block data as strings. ("BlockName" "Y" "X" "Z" [attributes...])
(defun DA:GetBlockData (BlockObject / Data Coordinates Attributes Entity)
  (setq Coordinates (cdr (assoc 10 (entget BlockObject))))
  (setq Data (list (nth 2 Coordinates) (nth 1 Coordinates) (nth 0 Coordinates) (cdr (assoc 2 (entget BlockObject)))))
  (while
    (and
      (null Attributes)
      (setq BlockObject (entnext BlockObject))
      (= "ATTRIB" (cdr (assoc 0 (setq Entity (entget BlockObject)))))
    )
    (setq Data (cons (cdr (assoc 1 Entity)) Data))
  )
  (reverse Data)
)

; Adds attribute columns to block data grid if grid has less attribute columns than Count.
; Count - [long] Count of needed attributes in block data grid.
(defun DA:UpdateGridAttr (Count /)
  (setq Count (- Count (DA:CurrentAttrCount)))
  (if (> Count 0)
    (while (/= Count 0)
      (setq Count (1- Count))
      (dcl-Grid-AddColumns BlockIE/Main/GridData (list (list "Attr" 1 50)))
    )
  )
)

; Returns [long] attribute column count in block data grid.
(defun DA:CurrentAttrCount ()
  (- (dcl-Grid-GetColumnCount BlockIE/Main/GridData) 4)
)

; Get configuration information from AutoCAD cfg file AppData section of BlockImportExport.
; Name - [str] Name of configuaration property.
; Returns: [str] value of configuration property or ""
(defun DA:GetCfgValue (Name / result)
  (if (not (setq result (getcfg (strcat "AppData/DA-BlockImportExport/" Name))))
	""
	result
  )
)

; Save configuration information to AutoCAD cfg file AppData section of BlockImportExport.
; Name  - [str] Name of configuaration property
; Value - [str] Value of configuaration property. The string can be up to 512 characters in length.
(defun DA:SetCfgValue (Name Value /)
  (setcfg (strcat "AppData/DA-BlockImportExport/" Name) Value)
)

