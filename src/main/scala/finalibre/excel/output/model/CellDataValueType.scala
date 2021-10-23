package finalibre.excel.output.model

enum CellDataValueType(val strName : String):
  case Integer extends CellDataValueType("int")
  case Double  extends CellDataValueType("double")
  case Date extends CellDataValueType("date")
  case DateTime extends CellDataValueType("datetime")
  case String extends CellDataValueType("string")

  def fromString(str : String) = CellDataValueType.values.find(_.toString == str).getOrElse(String)
  override def toString() = strName
