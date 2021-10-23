package finalibre.excel.output.model

case class FreezePaneSpecification(
                                    nrOfCols : Int,
                                    nrOfRows : Int,
                                    leftMostColumn : Option[Int] = None,
                                    topRow : Option[Int] = None
                                  )
