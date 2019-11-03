import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.Row
import java.lang.IllegalStateException

fun Row.getCell(c: Char): Cell? {
    val abc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    return this.getCell(abc.indexOf(c))
}

fun Cell.getValueAsPlainString(): String {
    val dataFormatter = DataFormatter()
    return dataFormatter.formatCellValue(this)
}

fun getValueFromCell(cell: Cell): Any? {

    return when (cell.cellType) {
        CellType._NONE -> throw IllegalStateException()
        CellType.NUMERIC -> cell.numericCellValue
        CellType.STRING -> cell.stringCellValue
        CellType.FORMULA -> cell.cellFormula
        CellType.BLANK -> null
        CellType.BOOLEAN -> cell.booleanCellValue
        CellType.ERROR -> throw IllegalStateException()
    }
}