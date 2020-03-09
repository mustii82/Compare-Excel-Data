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
private val filePathSource = desktopPath + "MC/2019-07-07_HC_Masterliste Auszug ATB.xlsx"
private val filePathMaster = desktopPath + "MC/HC_Masterliste_latest_shyla.xlsx"
private val startRow = 6
private val sourceLimitRow = 1230
private val masterLimitRow = 1337
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
                println(lineSeperator + "ERROR AT ROW: $rowNumber".black().redBackground() + "$lineSeperator$message".red())
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
                    printRowError(id + " seams to Exist already in ROW: ${rowMap[id].toString().black().redBackground()}" + lineSeperator + "This Employee is maybe included multiple times and this instance will be ignored".red())

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

fun main() {

    printOperation("START")
    val workbookMaster = XSSFWorkbook(FileInputStream(File(filePathMaster)))
    val workbookSource = XSSFWorkbook(FileInputStream(File(filePathSource)))

    printOperation("Get Data from Sheets")
    printOperation("Get Data from MasterSheet")
    val masterDataMap = getDataFromSheet(workbookMaster, masterLimitRow)
    printOperation("Get Data from SourceSheet")
    val sourceDataMap = getDataFromSheet(workbookSource, sourceLimitRow)

    printOperation("Get Data from Sheets Finished")
    println("MasterData Count: " + masterDataMap.count())
    println("SourceData Count: " + sourceDataMap.count())

    printOperation("Check Excel Sheets".toUpperCase())
    printOperation("Check Master Data")
    checkSheetData(masterDataMap)
    printOperation("Check Source Data")
    checkSheetData(sourceDataMap)

    printOperation("Start Comparison between SourceData and MasterData")
    val doesntExistList = ArrayList<Data>()
    var equalCounter = 0
    var differenceCounter = 0

    sourceDataMap.values.forEach { sourceData ->
        val dataFromMaster = masterDataMap[getIdentificator(sourceData)]

        if (dataFromMaster == null) {
            doesntExistList.add(sourceData)
        } else if (sourceData == dataFromMaster) {
            equalCounter += 1
            //println("EQUAL")
        } else {
            differenceCounter += 1
            printWarning("DIFFERENCES in ${getIdentificator(sourceData)}".yellow())

            printDifferences(sourceData, dataFromMaster)
            printCompareDataTable(sourceData, dataFromMaster, "SHEET", "SOURCE", "MASTER")

            lineSeperator.println()
        }
    }

    if (doesntExistList.isNotEmpty()) {
        printError("Following ${doesntExistList.size} SourceDataSets doesnt exist in MasterSheet")
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


fun printDifferences(sourceData: Data, masterData: Data) {

    val propNames: List<String> = Data::class.memberProperties.map { it.name }
    val masterValues: List<String> = Data::class.memberProperties.map { it.get(masterData) as String }
    val sourceValues: List<String> = Data::class.memberProperties.map { it.get(sourceData) as String }

    // Print Differences Warning
    for (i in propNames.indices) {
        val masterValue = masterValues[i]
        val sourceValue = sourceValues[i]
        if (masterValue != sourceValue)
            println(("There is a Difference in the Property: " + propNames[i]).red())
    }
}
