import kotlin.reflect.full.memberProperties

data class DataHolder(val row:Int,val data: Data)

data class Data(
    val `1surname`: String,
    val `2firstname`: String,
    val employeeType: String,
    val purpose: String,
    val reasonOfHiring: String,
    val company: String,
    val location: String,
    val department: String,
    val eMail: String,
    val userID: String,
    val dateOfEntry: String,
    val dateOfExits: String,
    val role: String,
    val workstream: String,
    val workpackage: String,
    val confirmation: String

) {
    fun isEmpty() = Data::class.memberProperties.all { (it.get(this) as String).isBlank() }
}