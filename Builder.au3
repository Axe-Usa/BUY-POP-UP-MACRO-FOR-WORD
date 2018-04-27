Local $oErrorHandler = ObjEvent("AutoIt.Error", "_ErrFunc")
$objWord = ObjCreate("Word.Application")
$objWord.Visible = True
$objDoc = $objWord.Documents.Add()
$objSelection = $objWord.Selection
Local $sMyPDF = @DesktopDir & "/run.xls"
$objSelection.InlineShapes.AddOLEObject("",$sMyPDF)
