import java.io.File

fun main(args: Array<String>) {
    File("./Result.xlsx").delete()
    ResultFile.writeFile()
}