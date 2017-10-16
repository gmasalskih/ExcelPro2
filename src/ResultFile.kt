import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileOutputStream
import java.io.IOException

object ResultFile {
    private val wbOut = XSSFWorkbook()
    private val OUT_FILE_NAME = "./Result.xlsx"

    init {
        File(OUT_FILE_NAME).delete()
    }

    private fun resultsOfCourses() {
        val sheetOut = wbOut.createSheet("resultsOfCourses")
        Excel.getResultsOfCourses().forEach { courseName, result ->
            val row = sheetOut.createRow(sheetOut.lastRowNum + 1)
            row.createCell(0).setCellValue(courseName)
            row.createCell(1).setCellType(CellType.NUMERIC)
            row.getCell(1).setCellValue(result.toDouble())
        }
    }

    private fun resultsOfGroups() {
        val sheetOut = wbOut.createSheet("resultsOfGroups")
        Excel.getResultsOfGroups().forEach { groupName, result ->
            val row = sheetOut.createRow(sheetOut.lastRowNum + 1)
            row.createCell(0).setCellValue(groupName)
            row.createCell(1).setCellType(CellType.NUMERIC)
            row.getCell(1).setCellValue(result.toDouble())
        }
    }

    private fun resultsOfGroupsByCourses() {
        val sheetOut = wbOut.createSheet("resultsOfGroupsByCourses")
        Excel.getResultsOfGroupsByCourses().forEach { groupName, coursesMap ->
            coursesMap.forEach { courseName, result ->
                val row = sheetOut.createRow(sheetOut.lastRowNum + 1)
                row.createCell(0).setCellValue(groupName)
                row.createCell(1).setCellValue(courseName)
                row.createCell(2).setCellType(CellType.NUMERIC)
                row.getCell(2).setCellValue(result.toDouble())
            }
        }
    }

    private fun resultsOfCoursesByGroups() {
        val sheetOut = wbOut.createSheet("resultsOfCoursesByGroups")
        Excel.getResultsOfCoursesByGroups().forEach { courseName, groupsMap ->
            groupsMap.forEach { groupName, result ->
                val row = sheetOut.createRow(sheetOut.lastRowNum + 1)
                row.createCell(0).setCellValue(courseName)
                row.createCell(1).setCellValue(groupName)
                row.createCell(2).setCellType(CellType.NUMERIC)
                row.getCell(2).setCellValue(result.toDouble())
            }
        }
    }

    private fun amountSuccessEmployeesInGroups() {
        val sheetOut = wbOut.createSheet("amountSuccessEmployeesInGroups")
        Excel.getAmountSuccessEmployeesInGroups().forEach { groupName, (amountAll, amountSuccess, res) ->
            val row = sheetOut.createRow(sheetOut.getLastRowNum() + 1)
            row.createCell(0).setCellValue(groupName)
            row.createCell(1).setCellValue(amountAll.toString())
            row.createCell(2).setCellValue(amountSuccess.toString())
            row.createCell(3).setCellType(CellType.NUMERIC)
            row.getCell(3).setCellValue(res.toDouble())
        }
    }

    fun writeFile() {
        resultsOfCourses()
        resultsOfGroups()
        amountSuccessEmployeesInGroups()
        resultsOfGroupsByCourses()
        resultsOfCoursesByGroups()
        try {
            FileOutputStream(OUT_FILE_NAME).use {
                wbOut.write(it)
                wbOut.close()
            }
        } catch (e: IOException) {
            e.printStackTrace()
        }
    }


}