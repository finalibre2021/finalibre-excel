package finalibre.excel.output.model

case class KeyValueData(header : Option[String], pairs : List[(String, CellData)]) extends SheetElement:
  override def spaceAfter = 2
