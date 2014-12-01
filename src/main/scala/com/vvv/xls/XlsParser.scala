package com.vvv.xls

import java.io.InputStream

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.poifs.filesystem.POIFSFileSystem
import org.apache.poi.ss.usermodel.{Cell, Row}

import scala.collection.JavaConverters._

class XlsParser {

  def parse(is: InputStream, sheetIndex: Int, columns: Column*): Seq[Map[Column, Any]] = {
    require(columns.nonEmpty, "Please specify a list of columns")

    val fs = new POIFSFileSystem(is)
    val wb = new HSSFWorkbook(fs)
    val sheet = wb.getSheetAt(sheetIndex)
    val rows = sheet.asScala

    parse(rows, columns)
  }

  private def parse(rows: Iterable[Row], columns: Seq[Column]): Seq[Map[Column, Any]] = {
    val indexColumnMapping = mapColumnsToIndexes(rows.head, columns)

    rows.tail.map { row =>
      (for (i <- 0 until row.getLastCellNum if indexColumnMapping.get(i).isDefined) yield
        indexColumnMapping(i) -> parse(row.getCell(i))).toMap
    }.toSeq
  }

  private def mapColumnsToIndexes(row: Row, columns: Seq[Column]): Map[Int, Column] =
    (for (i <- 0 until row.getLastCellNum;
          column <- matchColumn(row.getCell(i), columns)
    ) yield i -> column).toMap

  private def matchColumn(cell: Cell, columns: Seq[Column]): Option[Column] =
    if (cell == null) None
    else columns.find(_.name == cell.getStringCellValue)

  private def parse(cell: Cell): Any = {
    if (cell == null) None
    else cell.getCellType match {
      case Cell.CELL_TYPE_BOOLEAN => cell.getBooleanCellValue
      case Cell.CELL_TYPE_NUMERIC => cell.getNumericCellValue
      case Cell.CELL_TYPE_STRING => cell.getStringCellValue
      case _ => throw new RuntimeException(s"Unsupported cell type - ${cell.getCellType}. Cell - ${cell.toString}")
    }
  }

}
