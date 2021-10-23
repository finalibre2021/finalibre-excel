package finalibre.excel.output

import finalibre.excel.output.model.*
import org.apache.poi.ss.util.WorkbookUtil
import org.apache.poi.xssf.usermodel.*


object ExcelExporter:

  def writeTo(row: XSSFRow, colIndx: Int, cellData: CellData, isBold: Boolean)(implicit objs: CreationObjects) : Unit =
    writeTo(row, colIndx, Some(cellData), isBold)

  def writeTo(row: XSSFRow, colIndx: Int, cellDataOpt: Option[CellData], isBold: Boolean)(implicit objs: CreationObjects) : Unit =
    val cell = row.createCell(colIndx)
    cellDataOpt match
      case None => cell.setBlank()
      case Some(cellData) =>
        (cellData.vt, cellData.value) match
          case (CellDataValueType.Date, v: java.util.Date) =>
            cell.setCellValue(v)
            cell.setCellStyle(objs.dateStyle)

          case (CellDataValueType.DateTime, v: java.util.Date) =>
            cell.setCellValue(v)
            cell.setCellStyle(objs.dateTimeStyle)

          case (CellDataValueType.Double, d: Double) =>
            cell.setCellValue(d)

          case (_, v) =>
            cell.setCellValue(v.toString)
            if (isBold) cell.setCellStyle(objs.boldStringStyle)









  def stringCell(str: String) = CellData(str, CellDataValueType.String) : CellData

  case class CreationObjects(wb: XSSFWorkbook):
    val helper = wb.getCreationHelper
    val dateStyle = {
      val style = wb.createCellStyle
      style.setDataFormat(helper.createDataFormat.getFormat("dd-MM-yyyy"))
      style
    }

    val dateTimeStyle = {
      val style = wb.createCellStyle
      style.setDataFormat(helper.createDataFormat.getFormat("dd-MM-yyyy HH:mm:ss"))
      style
    }

    val boldStringStyle = {
      val style = wb.createCellStyle
      val font = wb.createFont
      font.setBold(true)
      style.setFont(font)
      style
    }



