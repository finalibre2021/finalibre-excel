package finalibre.excel.output

import finalibre.excel.output.ExcelExporter.Exportables.{MapExportable, ProductExportable}
import finalibre.excel.output.model.*
import org.apache.poi.ss.util.WorkbookUtil
import org.apache.poi.xssf.usermodel.*
import org.slf4j.LoggerFactory

import java.io.ByteArrayOutputStream
import scala.util.Try


object ExcelExporter:

  private val logger = LoggerFactory.getLogger(this.getClass)

  def exportCollections(colls : Seq[(String, Seq[Product], Option[FreezePaneSpecification])], headerMap : Map[String, String] = Map()) =
    val generalised = colls.map(col => col.copy(_2 = Exportables.ProductExportable(col._2)))
    exportExportables(generalised, headerMap)


  def exportExportables(colls : Seq[(String, Exportables.Exportable, Option[FreezePaneSpecification])], headerMap : Map[String, String] = Map()) =
    val mapped = (colls.toList.map {
      case (sheetName, ProductExportable(coll), fpSpec) => (sheetName, Reflective.convert(coll, headerMap), fpSpec)
      case (sheetName, MapExportable(headers, maps), fpSpec) => {
        val coll = transform(headers, maps)
        (sheetName, Reflective.convert(coll, headerMap), fpSpec)
      }

    }).collect {
      case (sheetName, Some(tabDat), fpSpec) => SheetData(sheetName, false, List(tabDat), fpSpec)
    }
    exportSheets(mapped)


  def exportSheets(sheets: List[SheetData]) =
    Try {
      val wb = new XSSFWorkbook
      implicit val objs = CreationObjects(wb)
      sheets.zipWithIndex.foreach {
        case (sheetData, indx) => {
          exportSheet(sheetData, indx, wb)
        }
      }
      sheets.zipWithIndex.find(_._1.isSelected).foreach {
        case (selSheet, indx) => {
          wb.setActiveSheet(indx)
          wb.setSelectedTab(indx)
        }
      }

      val outputStream = new ByteArrayOutputStream
      wb.write(outputStream)
      outputStream.close()
      outputStream.toByteArray()
    }


  def exportSheet(sheetData: SheetData, indx: Int, wb: XSSFWorkbook)(implicit objs: CreationObjects) =
    val sheetName = WorkbookUtil.createSafeSheetName(if (sheetData.sheetName.length > 30) sheetData.sheetName.substring(0, 30) else sheetData.sheetName)
    val sheet = wb.createSheet(sheetName)
    var currentIndx = 0
    def recurse(elems: List[SheetElement]) : Unit =
      def doElem(el: SheetElement) = {
        el match {
          case td: TableData    => exportTable(td, currentIndx, sheet)
          case kv: KeyValueData => exportKeyValue(kv, currentIndx, sheet)
          case _                => currentIndx
        }
      }
      elems match
        case Nil =>
        case el :: Nil =>
          doElem(el)

        case el :: lis =>
          currentIndx = doElem(el)
          currentIndx += 1
          for (i <- 0 until el.spaceAfter) {
            sheet.createRow(currentIndx)
            recurse(lis)
          }



    recurse(sheetData.elements)
    sheetData.freezePaneSpecification match
      case Some(FreezePaneSpecification(nrOfCols, nrOfRows, Some(leftMost), Some(topRow))) => sheet.createFreezePane(nrOfCols, nrOfRows, leftMost, topRow)
      case Some(FreezePaneSpecification(nrOfCols, nrOfRows, _, _)) => sheet.createFreezePane(nrOfCols, nrOfRows)
      case _ =>




  def exportTable(tableData: TableData, insertRowIndx: Int, sheet: XSSFSheet)(implicit objs: CreationObjects) =
    val header = tableData.headers
    var currentRowIndx = insertRowIndx
    val headerRow = sheet.createRow(currentRowIndx)
    header.zipWithIndex.foreach {
      case (header, indx) =>
        writeTo(headerRow, indx, stringCell(header.showName), true)

    }
    tableData.rowData.foreach {
      case rDat =>
        currentRowIndx += 1
        val row = sheet.createRow(currentRowIndx)
        header.zipWithIndex.foreach {
          case (h, indx) =>
            writeTo(row, indx, rDat.dataMap.get(h.id), rDat.isBold)
        }
    }
    header.zipWithIndex.foreach(p => sheet.autoSizeColumn(p._2))
    currentRowIndx


  def exportKeyValue(keyVal: KeyValueData, insertRowIndx: Int, sheet: XSSFSheet)(implicit objs: CreationObjects) =
    var currentRowIndex = insertRowIndx
    keyVal.header match {
      case None =>
      case Some(h) => {
        val hRow = sheet.createRow(currentRowIndex)
        writeTo(hRow, 0, stringCell(h), true)
        currentRowIndex += 1
      }
    }
    keyVal.pairs.foreach {
      case (k, v) => {
        val row = sheet.createRow(currentRowIndex)
        writeTo(row, 0, stringCell(k), false)
        writeTo(row, 1, v, false)
        currentRowIndex += 1
      }
    }
    currentRowIndex - 1


  def writeTo(row: XSSFRow, colIndx: Int, cellData: CellData, isBold: Boolean)(implicit objs: CreationObjects) : Unit =  { writeTo(row, colIndx, Some(cellData), isBold) }

  def writeTo(row: XSSFRow, colIndx: Int, cellDataOpt: Option[CellData], isBold: Boolean)(implicit objs: CreationObjects) : Unit =
    val cell = row.createCell(colIndx)
    cellDataOpt match
      case None => cell.setBlank
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
            cell.setCellStyle(objs.doubleStyle)
          case (CellDataValueType.Integer, i: Int) =>
            cell.setCellValue(i)

          case (CellDataValueType.Integer, l: Long) =>
            cell.setCellValue(l)
            cell.setCellStyle(objs.longStyle)


          case (_, v) =>
            cell.setCellValue(v.toString)
            if (isBold) cell.setCellStyle(objs.boldStringStyle)







  def stringCell(str: String) = CellData(str, CellDataValueType.String)

  private def transform(headers : Seq[String], maps : Seq[Map[String, Any]]) : Seq[Product] =
    val headerLength = headers.size
    maps.map {
      case map => new Product {
        override def productArity: Int = headerLength
        override def productElement(n: Int): Any = map.get(headers(n))
        override def canEqual(that: Any): Boolean = false
        override def productElementNames: Iterator[String] = headers.iterator
      }
    }

  object Exportables:
    sealed trait Exportable
    case class ProductExportable(products : Seq[Product]) extends Exportable
    case class MapExportable(headers : Seq[String], maps : Seq[Map[String, Any]]) extends Exportable



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
    val doubleStyle = {
      val style = wb.createCellStyle()
      style.setDataFormat(helper.createDataFormat().getFormat("###,##0.00"))
      style
    }
    val longStyle = {
      val style = wb.createCellStyle()
      style.setDataFormat(helper.createDataFormat().getFormat("#############0"))
      style
    }
    val boldStringStyle = {
      val style = wb.createCellStyle
      val font = wb.createFont
      font.setBold(true)
      style.setFont(font)
      style
    }





