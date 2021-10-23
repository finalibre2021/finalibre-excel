package finalibre.excel.output

import java.sql.{Date, Timestamp}
import finalibre.excel.output.model.*


protected object Reflective {
  val vt = CellDataValueType

  def convert[A <: Product](seq : Seq[A], headerMap : Map[String, String] = Map()) = {
    seq.headOption match {
      case None => None
      case Some(head) => {
        val longestHeaders = seq.map(en => en.productElementNames.toList).maxBy(_.size)
        val attrs = longestHeaders.toList.zipWithIndex
        val headers = attrs.map {
          case (attr,indx) => TableHeaderData(attr, headerMap.getOrElse(attr, attr))
        }
        val rows = seq.map {
          case entry => RowData((attrs.map {
            case (nam, indx) => nam -> (entry.productElement(indx) match {
              case null => CellData("", vt.String)
              case None => CellData("", vt.String)
              case d : Double => CellData(d, vt.Double)
              case Some(d : Double) => CellData(d, vt.Double)
              case i : Int => CellData(i, vt.Integer)
              case Some(i : Int) => CellData(i, vt.Integer)
              case l : Long => CellData(l, vt.Integer)
              case Some(l : Long) => CellData(l, vt.Integer)
              case d : Date => CellData(d, vt.Date)
              case Some(d : Date) => CellData(d, vt.Date)
              case t : Timestamp => CellData(t, vt.DateTime)
              case Some(t : Timestamp) => CellData(t, vt.DateTime)
              case Some(oth) => CellData(oth.toString, vt.String)
              case oth => CellData(oth.toString, vt.String)
            })
          }).toMap, isBold = false)
        }
        Some(TableData(headers, rows.toList))

      }
    }
  }

}
