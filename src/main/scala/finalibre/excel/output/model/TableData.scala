package finalibre.excel.output.model

case class TableData (headers : List[TableHeaderData], rowData : List[RowData]) extends SheetElement:
  override def spaceAfter = 3
