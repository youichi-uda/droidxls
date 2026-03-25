package com.droidoffice.xls.sample

import android.os.Bundle
import android.util.Log
import android.widget.Button
import android.widget.TextView
import androidx.appcompat.app.AppCompatActivity
import kotlinx.coroutines.*

class MainActivity : AppCompatActivity() {

    private lateinit var tvLog: TextView
    private val scope = CoroutineScope(Dispatchers.Main + SupervisorJob())

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        tvLog = findViewById(R.id.tvLog)

        val demos = Demos(this)

        findViewById<Button>(R.id.btnBasicReadWrite).setOnClickListener { runDemo("Basic Read/Write") { demos.basicReadWrite() } }
        findViewById<Button>(R.id.btnStyles).setOnClickListener { runDemo("Styles") { demos.styles() } }
        findViewById<Button>(R.id.btnFormulas).setOnClickListener { runDemo("Formulas") { demos.formulas() } }
        findViewById<Button>(R.id.btnSheetOps).setOnClickListener { runDemo("Sheet Ops") { demos.sheetOperations() } }
        findViewById<Button>(R.id.btnDataFeatures).setOnClickListener { runDemo("Data Features") { demos.dataFeatures() } }
        findViewById<Button>(R.id.btnPassword).setOnClickListener { runDemo("Password") { demos.passwordProtection() } }
        findViewById<Button>(R.id.btnConvert).setOnClickListener { runDemo("Convert") { demos.csvHtmlExport() } }
        findViewById<Button>(R.id.btnFullReport).setOnClickListener { runDemo("Full Report") { demos.fullReport() } }
        findViewById<Button>(R.id.btnRunAll).setOnClickListener { runAllDemos(demos) }

        // Auto-run if launched with "autorun" extra
        if (intent?.getBooleanExtra("autorun", false) == true) {
            runAllDemos(demos)
        }
    }

    private fun runDemo(name: String, block: suspend () -> String) {
        scope.launch {
            log("--- $name ---")
            try {
                val result = withContext(Dispatchers.IO) { block() }
                log(result)
                log("OK\n")
            } catch (e: Exception) {
                log("FAILED: ${e.message}\n")
            }
        }
    }

    private fun runAllDemos(demos: Demos) {
        scope.launch {
            tvLog.text = ""
            val allDemos = listOf(
                "Basic Read/Write" to suspend { demos.basicReadWrite() },
                "Styles" to suspend { demos.styles() },
                "Formulas" to suspend { demos.formulas() },
                "Sheet Ops" to suspend { demos.sheetOperations() },
                "Data Features" to suspend { demos.dataFeatures() },
                "Password" to suspend { demos.passwordProtection() },
                "CSV/HTML" to suspend { demos.csvHtmlExport() },
                "Full Report" to suspend { demos.fullReport() },
            )
            var passed = 0
            var failed = 0
            for ((name, block) in allDemos) {
                log("--- $name ---")
                try {
                    val result = withContext(Dispatchers.IO) { block() }
                    log(result)
                    log("OK\n")
                    passed++
                } catch (e: Exception) {
                    log("FAILED: ${e.message}\n")
                    failed++
                }
            }
            log("=============================")
            log("Results: $passed passed, $failed failed")
        }
    }

    private fun log(msg: String) {
        tvLog.append("$msg\n")
        Log.d(TAG, msg)
    }

    companion object {
        private const val TAG = "DroidXLS"
    }

    override fun onDestroy() {
        super.onDestroy()
        scope.cancel()
    }
}
