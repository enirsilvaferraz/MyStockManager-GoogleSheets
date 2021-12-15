package com.ferraz.sheets_communication

import android.content.Context
import android.content.Intent
import android.os.Bundle
import androidx.appcompat.app.AppCompatActivity
import com.google.android.gms.auth.api.signin.GoogleSignIn
import com.google.android.gms.auth.api.signin.GoogleSignInOptions
import com.google.android.gms.common.api.Scope
import com.google.api.client.extensions.android.http.AndroidHttp
import com.google.api.client.googleapis.extensions.android.gms.auth.GoogleAccountCredential
import com.google.api.client.json.jackson2.JacksonFactory
import com.google.api.services.sheets.v4.Sheets
import com.google.api.services.sheets.v4.SheetsScopes
import com.google.api.services.sheets.v4.model.BatchUpdateValuesRequest
import com.google.api.services.sheets.v4.model.Spreadsheet
import com.google.api.services.sheets.v4.model.SpreadsheetProperties
import com.google.api.services.sheets.v4.model.ValueRange
import kotlinx.coroutines.CoroutineScope
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.MainScope
import kotlinx.coroutines.cancel
import kotlinx.coroutines.launch
import timber.log.Timber

class TestSheetsApiActivity : AppCompatActivity(), CoroutineScope by MainScope() {

    companion object {
        private const val REQUEST_SIGN_IN = 1
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.testsheetsapi)

        requestSignIn(this)
    }

    override fun onDestroy() {
        super.onDestroy()
        cancel()
    }

    override fun onActivityResult(requestCode: Int, resultCode: Int, data: Intent?) {
        super.onActivityResult(requestCode, resultCode, data)

        if (requestCode == REQUEST_SIGN_IN) {
            if (resultCode == RESULT_OK) {
                GoogleSignIn.getSignedInAccountFromIntent(data)
                    .addOnSuccessListener { account ->
                        val scopes = listOf(SheetsScopes.SPREADSHEETS)
                        val credential = GoogleAccountCredential.usingOAuth2(this, scopes)
                        credential.selectedAccount = account.account

                        val jsonFactory = JacksonFactory.getDefaultInstance()
                        // GoogleNetHttpTransport.newTrustedTransport()
                        val httpTransport =  AndroidHttp.newCompatibleTransport()
                        val service = Sheets.Builder(httpTransport, jsonFactory, credential)
                            .setApplicationName(getString(R.string.app_name))
                            .build()

                        createSpreadsheet(service)
                    }
                    .addOnFailureListener { e ->
                        Timber.e(e)
                    }
            }
        }
    }

    private fun requestSignIn(context: Context) {
        /*
        GoogleSignIn.getLastSignedInAccount(context)?.also { account ->
            Timber.d("account=${account.displayName}")
        }
         */

        val signInOptions = GoogleSignInOptions.Builder(GoogleSignInOptions.DEFAULT_SIGN_IN)
             .requestEmail()
            // .requestScopes(Scope(SheetsScopes.SPREADSHEETS_READONLY))
            .requestScopes(Scope(SheetsScopes.SPREADSHEETS))
            .build()
        val client = GoogleSignIn.getClient(context, signInOptions)

        startActivityForResult(client.signInIntent, REQUEST_SIGN_IN)
    }

    private fun createSpreadsheet(service: Sheets) {
        var spreadsheet = Spreadsheet()
            .setProperties(
                SpreadsheetProperties()
                    .setTitle("CreateNewSpreadsheet")
            )

        launch(Dispatchers.Default) {
            try {

                val sheetName = "Página1"
                val spreadsheetId = "1qcVI29gAQ1C3oy8-1Bvj5EWN1HBKeryUpwjYQqCSl_U"

                appendRow(sheetName, service, spreadsheetId)

                //updateRow(service, spreadsheetId, sheetName)


            } catch (e: Exception) {
                Timber.e(e)
            }
        }
    }

    private fun appendRow(sheetName: String, service: Sheets, spreadsheetId: String) {
        val rows = listOf(
            listOf<String>(),
            listOf("Col 01", "Col 02", "Col 03"),
            listOf("Col 01", "Col 02", "Col 03")
        )

        val body = ValueRange().setValues(rows)

        val range = "'$sheetName'!A1"

        val result = service.spreadsheets().values().append(spreadsheetId, range, body)
            .setValueInputOption("RAW")
            .setInsertDataOption("INSERT_ROWS")
            .execute()

        Timber.d("newRows=${result.updates.updatedRows}")
    }

    private fun updateRow(service: Sheets, spreadsheetId: String, sheetName: String) {



        val rows = listOf<ValueRange>(
            ValueRange()
                .setRange("'$sheetName'!A1:B3")
                .setValues(
                    listOf(
                        listOf<Any>("Update A1...", "X1"),
                        listOf<Any>("Update A2..."),
                        listOf<Any>("Update A3...", "X3"),
                    )
                ),
            /*
            ValueRange()
                .setRange("'$sheetName'!B2")
                .setValues(
                    listOf(
                        listOf<Any>("Updaate B2...")
                    )
                )
                    */
        )

        val body = BatchUpdateValuesRequest()
            .setValueInputOption("RAW")
            .setData(rows)

        val result = service.spreadsheets().values().batchUpdate(spreadsheetId, body).execute()

        Timber.d("updatedRows=${result.totalUpdatedRows}")
    }

    private fun createSheet(spreadsheet: Spreadsheet, service: Sheets) {
        var spreadsheet1 = spreadsheet
        spreadsheet1 = service.spreadsheets().create(spreadsheet1).execute()
        Timber.d("ID: ${spreadsheet1.spreadsheetId}")
    }
}