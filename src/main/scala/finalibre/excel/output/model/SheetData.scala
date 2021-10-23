package finalibre.excel.output.model

case class SheetData(sheetName : String, isSelected : Boolean, elements : List[SheetElement])