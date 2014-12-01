package com.vvv.xls

import org.scalatest.FlatSpec
import org.scalatest.Matchers._

class XlsParserTest extends FlatSpec {

  behavior of classOf[XlsParser].getSimpleName

  it should "parse xls file" in {
    val xlsParser = new XlsParser
    val rows = xlsParser.parse(getClass.getClassLoader.getResourceAsStream("xls.xls"), 0, A, B, C)
    rows should have size 2
    println(rows)
    rows should equal(List(
      Map(A -> 1, B -> "b", C -> "c"),
      Map(A -> 2, B -> None, C -> "c2")
    ))
  }

  private case object A extends Column("a")

  private case object B extends Column("b")

  private case object C extends Column("c")

}
