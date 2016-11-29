package io.terrafino.emule

import java.io.FileInputStream

import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.WorkbookFactory

import io.terrafino.logging.Logger

object ExcelMule extends App with Logger {

  logger.info("Start")

  val fis = new FileInputStream("test.xlsx")

  val workbook = WorkbookFactory.create(fis)
  val evaluator = this.workbook.getCreationHelper().createFormulaEvaluator()
  val formatter = new DataFormatter(true)

  val sheet = workbook.getSheetAt(0)

  if (sheet.getPhysicalNumberOfRows() > 0) {
    for (i <- 0 to sheet.getLastRowNum) {
      val row = sheet.getRow(i)
      if (row != null) {
        for (j <- 0 to row.getLastCellNum) {
          val cell = row.getCell(j)
          if (cell == null) {
            println("")
          } else {
            if (cell.getCellTypeEnum() != CellType.FORMULA) {
              println(formatter.formatCellValue(cell))
            } else {
              println("f: "+formatter.formatCellValue(cell, this.evaluator))
            }
          }
        }
      }
    }
  }

  fis.close()

  logger.info("End")
}