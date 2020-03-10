import com.andreapivetta.kolor.*
import com.jakewharton.fliptables.FlipTable
import mkl.extensions.types.println
import mkl.global.variables.desktopPath
import mkl.global.variables.lineSeperator
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import kotlin.reflect.full.memberProperties

// SETTINGS
private var filePathBase = desktopPath + "src.xlsx"
private var filePathChange = desktopPath + "change"
private val startRow = 6
private val changeLimitRow = 1230
private val baseLimitRow = 1337
fun getIdentificator(data: Data) = data.`2firstname` + " " + data.`1surname`


fun printOperation(message: String) =
    println(lineSeperator + ("-".repeat(10) + message + "-".repeat(10)).black().magentaBackground())

fun printError(message: String) = println("ERROR:".black().redBackground() + " " + message.red())
fun printWarning(message: String) = println("WARNING:".black().yellowBackground() + " " + message.yellow())

fun getDataFromSheet(workbook: XSSFWorkbook, limit: Int): HashMap<String, Data> {

    val rows = workbook.getSheetAt(0).iterator()
    val dataMap = HashMap<String, Data>()

    val rowMap = HashMap<String, Int>()

    rows.forEach { row ->
        val rowNumber = row.rowNum + 1

        if (rowNumber in startRow..limit) {
            // println("ROW: $rowNumber")

            fun printRowError(message: String) {
                println(
                    lineSeperator + "ERROR AT ROW: $rowNumber".black().redBackground() + "$lineSeperator$message".red()
                )
            }

            val data = Data(
                row.getCell('E')!!.getValueAsPlainString(),
                row.getCell('F')!!.getValueAsPlainString(),
                row.getCell('G')!!.getValueAsPlainString(),
                row.getCell('H')!!.getValueAsPlainString(),
                row.getCell('I')!!.getValueAsPlainString(),
                row.getCell('J')!!.getValueAsPlainString(),
                row.getCell('K')!!.getValueAsPlainString(),

                row.getCell('O')!!.getValueAsPlainString(),
                row.getCell('P')!!.getValueAsPlainString(),
                row.getCell('Q')!!.getValueAsPlainString(),
                row.getCell('R')!!.getValueAsPlainString(),
                row.getCell('S')!!.getValueAsPlainString(),
                row.getCell('T')!!.getValueAsPlainString(),
                row.getCell('U')!!.getValueAsPlainString(),
                row.getCell('V')!!.getValueAsPlainString(),
                row.getCell('W')!!.getValueAsPlainString()
            )


            //check(!(data.firstname.isBlank() || data.surname.isBlank())) { "NO IDENTIFICATION POSSIBLE LATER (MISSING KEY HERE)" }

            if (data.isEmpty()) { // Complete Row is Empty
                //printError("Following Dataset is Empty and gonna be ignored:")
                //data.toJson().println()
            } else if (data.`2firstname`.isBlank() || data.`1surname`.isBlank()) {
                printRowError("NO IDENTIFICATION POSSIBLE LATER (MISSING KEY HERE Surname + Firstname): The following Dataset will be ignored:")
                println(getDataTable(data).red())
            } else { // Add to DataMap
                //rowNumber.println()
                //data.println()
                val id = getIdentificator(data)
                if (dataMap.containsKey(id)) {
                    printRowError(
                        id + " seams to Exist already in ROW: ${rowMap[id].toString().black()
                            .redBackground()}" + lineSeperator + "This Employee is maybe included multiple times and this instance will be ignored".red()
                    )

                    printCompareDataTable(
                        dataMap[id]!!,
                        data,
                        "VERSION",
                        "ALREADY INCLUDED FROM ROW: ${rowMap[id]}",
                        "IGNORED FROM ROW: $rowNumber"
                    )

                } else {
                    dataMap[id] = data
                    rowMap[id] = rowNumber
                }
            }
        }

    }

    return dataMap
}

fun main(args: Array<String>) {

    if (args.size == 2) {
        filePathBase = args[0]
        filePathChange = args[1]
    }

    printOperation("START")
    val workbookBase = XSSFWorkbook(FileInputStream(File(filePathBase)))
    val workbookChange = XSSFWorkbook(FileInputStream(File(filePathChange)))

    printOperation("Get Data from Sheets")
    printOperation("Get Data from BaseSheet")
    val baseDataMap = getDataFromSheet(workbookBase, baseLimitRow)
    printOperation("Get Data from ChangedSheet")
    val changeDataMap = getDataFromSheet(workbookChange, changeLimitRow)

    printOperation("Get Data from Sheets Finished")
    println("BaseData Count: " + baseDataMap.count())
    println("ChangeData Count: " + changeDataMap.count())

    printOperation("Check Excel Sheets".toUpperCase())
    printOperation("Check Base Data")
    checkSheetData(baseDataMap)
    printOperation("Check Change Data")
    checkSheetData(changeDataMap)

    printOperation("Start Comparison between BaseData and ChangeData")
    val doesntExistList = ArrayList<Data>()
    var equalCounter = 0
    var differenceCounter = 0

    changeDataMap.values.forEach { changeData ->
        val dataFromBase = baseDataMap[getIdentificator(changeData)]

        if (dataFromBase == null) {
            doesntExistList.add(changeData)
        } else if (changeData == dataFromBase) {
            equalCounter += 1
            //println("EQUAL")
        } else {
            differenceCounter += 1
            printWarning("DIFFERENCES in ${getIdentificator(changeData)}".yellow())

            printDifferences(dataFromBase, changeData)
            printCompareDataTable(changeData, dataFromBase, "SHEET", "CHANGE", "BASE")

            lineSeperator.println()
        }
    }

    if (doesntExistList.isNotEmpty()) {
        printError("Following ${doesntExistList.size} SourceDataSets doesnt exist in BaseSheet")
        doesntExistList.forEach { println(getDataTable(it).red()) }
    }

    printOperation("FINISH")
    println(
        """
        doesntExistCounter: ${doesntExistList.size}
        equalCounter: $equalCounter
        differenceCounter: $differenceCounter
    """.trimIndent().red()
    )

}

fun checkSheetData(hashMap: HashMap<String, Data>) {
    fun printCheckPassed(checkPassed: Boolean, message: String) {
        if (checkPassed) println("Check Passed:".black().greenBackground() + " " + message.green())
    }

    // Check if UserID is uniqe
    var checkPassed = true
    val uniqeUserIDs = hashMap.values.map { it.userID.toUpperCase() }.distinct()
    uniqeUserIDs.forEach { userID ->
        val userIDCount = hashMap.values.count { !it.userID.isBlank() && it.userID.equals(userID, ignoreCase = true) }
        if (userIDCount > 1) {
            checkPassed = false
            printError("UserID: $userID appears $userIDCount times")
        }
    }
    printCheckPassed(checkPassed, "All UserIDs appear 1 time")

    // Check if eMail is uniqe
    checkPassed = true
    val uniqeEMails = hashMap.values.map { it.eMail.toLowerCase() }.distinct()
    uniqeEMails.forEach { eMail ->
        val eMailCount = hashMap.values.count { !it.eMail.isBlank() && it.eMail.equals(eMail, ignoreCase = true) }
        if (eMailCount > 1) {
            checkPassed = false
            printError("E-Mail: $eMail appears $eMailCount times")
        }
    }
    printCheckPassed(checkPassed, "All E-Mails appear 1 time")
}

fun getDataTable(data: Data): String {
    val propNames: List<String> = Data::class.memberProperties.map { it.name }
    val dataValues: List<String> = Data::class.memberProperties.map { it.get(data) as String }

    val headers = propNames.toTypedArray()
    val dataArray = arrayOf(
        dataValues.toTypedArray()
    )

    return FlipTable.of(headers, dataArray)
}


fun printCompareDataTable(data1: Data, data2: Data, header: String, data1Title: String, data2Title: String) {
    val propNames: List<String> = Data::class.memberProperties.map { it.name }
    val data1Values: List<String> = Data::class.memberProperties.map { it.get(data1) as String }
    val dat2Values: List<String> = Data::class.memberProperties.map { it.get(data2) as String }

    val headers = arrayOf(header) + propNames.toTypedArray()
    val data = arrayOf(
        arrayOf(data1Title) + data1Values.toTypedArray(),
        arrayOf(data2Title) + dat2Values.toTypedArray()
    )

    println(FlipTable.of(headers, data))
}


fun printDifferences(baseData: Data, changeData: Data) {

    val propNames: List<String> = Data::class.memberProperties.map { it.name }
    val baseValues: List<String> = Data::class.memberProperties.map { it.get(baseData) as String }
    val changeValues: List<String> = Data::class.memberProperties.map { it.get(changeData) as String }

    // Print Differences Warning
    for (i in propNames.indices) {
        if (baseValues[i] != changeValues[i])
            println(("There is a Difference in the Property: " + propNames[i]).red())
    }
}
