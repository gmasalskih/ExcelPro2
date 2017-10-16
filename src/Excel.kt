import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.io.IOException

object Excel {

    private val rows: Set<Row>

    init {
        rows = File("./").listFiles()
                .filter { it.absolutePath.contains(".xlsx") }
                .map { getSheet(it) }
                .flatMap { it?.asIterable()!! }
                .map { it as XSSFRow }
                .filter { it.rowNum > 1 }
                .map { getRow(it) }
                .toHashSet()
    }

    fun getResultsOfGroups(): Map<String, Float> {
        return rows.groupByTo(mutableMapOf(), { it.groupName })
                .mapValues { it.value.filter { it.isPass }.count().toFloat() / it.value.size.toFloat() }
    }

    fun getResultsOfCourses(): Map<String, Float> {
        return rows.groupByTo(mutableMapOf(), { it.courseName })
                .mapValues { it.value.filter { it.isPass }.count().toFloat() / it.value.size.toFloat() }
    }

    fun getAmountSuccessEmployeesInGroups(): Map<String, Triple<Int, Int, Float>> {
        return rows.groupByTo(mutableMapOf(), { it.groupName })
                .mapValues {
                    var all: Int = it.value.groupBy { it.id }.keys.size
                    var set = HashSet<Int>()
                    it.value.groupBy { it.id }.forEach { id, row ->
                        if (row.all { it.isPass }) set.add(id)
                    }
                    var pass = set.size
                    Triple(all, pass, pass.toFloat() / all.toFloat())
                }
    }

    fun getResultsOfGroupsByCourses(): Map<String, Map<String, Float>> {
        return rows.groupByTo(mutableMapOf(), { it.groupName })
                .mapValues {
                    var map = HashMap<String, Float>()
                    it.value.groupBy { it.courseName }
                            .forEach { k, v ->
                                map.put(k, v.filter { it.isPass }.count().toFloat() / v.size.toFloat())
                            }
                    map
                }
    }

    fun getResultsOfCoursesByGroups(): Map<String, Map<String, Float>> {
        return rows.groupByTo(mutableMapOf(), { it.courseName })
                .mapValues {
                    var map = HashMap<String, Float>()
                    it.value.groupBy { it.groupName }
                            .forEach { k, v ->
                                map.put(k, v.filter { it.isPass }.count().toFloat() / v.size.toFloat())
                            }
                    map
                }
    }

    // ---------- Other ----------

    private data class Row(val lastName: String,
                           val firstName: String,
                           val patronymic: String?,
                           val id: Int,
                           val courseName: String,
                           val isPass: Boolean,
                           val groupName: String)

    private enum class Index(val index: Int) {
        LAST_NAME(0),
        FIRST_NAME(1),
        PATRONYMIC(2),
        ID(3),
        COURSE_NAME(4),
        RESULT(5),
        GROUP(6)
    }

    private fun getRow(row: XSSFRow): Row {
        return Row(row.getCell(Index.LAST_NAME.index).stringCellValue.trim(),
                row.getCell(Index.FIRST_NAME.index).stringCellValue.trim(),
                row.getCell(Index.PATRONYMIC.index).stringCellValue.trim(),
                row.getCell(Index.ID.index).stringCellValue.trim().toInt(),
                row.getCell(Index.COURSE_NAME.index).stringCellValue.trim(),
                row.getCell(Index.RESULT.index).stringCellValue.trim() == "Завершенo",
                row.getCell(Index.GROUP.index).stringCellValue.trim())
    }

    private fun getSheet(file: File): XSSFSheet? {
        return try {
            FileInputStream(file).use { fis -> return XSSFWorkbook(fis).getSheetAt(0) }
        } catch (e: IOException) {
            e.printStackTrace()
            null
        }
    }
}