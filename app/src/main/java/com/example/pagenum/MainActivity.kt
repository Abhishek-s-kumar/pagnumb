package com.example.pptxprocessor

import android.content.Intent
import android.net.Uri
import android.os.Bundle
import android.view.View
import android.widget.Toast
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import androidx.lifecycle.lifecycleScope
import kotlinx.coroutines.launch
import java.text.SimpleDateFormat
import java.util.*
import com.example.pptxprocessor.databinding.ActivityMainBinding

class MainActivity : AppCompatActivity() {

    private lateinit var binding: ActivityMainBinding
    private var selectedFileUri: Uri? = null
    private lateinit var pptxProcessor: PPTXProcessor

    // File picker for input file
    private val pickInputFile = registerForActivityResult(
        ActivityResultContracts.GetContent()
    ) { uri ->
        if (uri != null) {
            selectedFileUri = uri
            val fileName = getFileName(uri)
            binding.tvSelectedFile.text = "Selected: $fileName"
            binding.btnProcessFile.isEnabled = true
            addLog("Selected file: $fileName")
        }
    }

    // File picker for output location
    private val createOutputFile = registerForActivityResult(
        ActivityResultContracts.CreateDocument("application/vnd.openxmlformats-officedocument.presentationml.presentation")
    ) { uri ->
        if (uri != null && selectedFileUri != null) {
            processFile(selectedFileUri!!, uri)
        }
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        binding = ActivityMainBinding.inflate(layoutInflater)
        setContentView(binding.root)

        pptxProcessor = PPTXProcessor(this)
        setupViews()
    }

    private fun setupViews() {
        binding.btnSelectFile.setOnClickListener {
            pickInputFile.launch("application/vnd.openxmlformats-officedocument.presentationml.presentation")
        }

        binding.btnProcessFile.setOnClickListener {
            if (selectedFileUri != null) {
                val timestamp = SimpleDateFormat("yyyyMMdd_HHmmss", Locale.getDefault()).format(Date())
                val outputFileName = "numbered_presentation_$timestamp.pptx"
                createOutputFile.launch(outputFileName)
            }
        }
    }

    private fun processFile(inputUri: Uri, outputUri: Uri) {
        binding.progressBar.visibility = View.VISIBLE
        binding.btnProcessFile.isEnabled = false
        binding.btnSelectFile.isEnabled = false

        lifecycleScope.launch {
            pptxProcessor.processPresentation(
                inputUri,
                outputUri,
                object : PPTXProcessor.ProcessingListener {
                    override fun onProgress(progress: Int, message: String) {
                        runOnUiThread {
                            binding.tvStatus.text = message
                            addLog(message)
                        }
                    }

                    override fun onComplete(success: Boolean, message: String) {
                        runOnUiThread {
                            binding.progressBar.visibility = View.GONE
                            binding.btnProcessFile.isEnabled = true
                            binding.btnSelectFile.isEnabled = true

                            if (success) {
                                binding.tvStatus.text = "✅ $message"
                                Toast.makeText(this@MainActivity, "Success! File saved.", Toast.LENGTH_LONG).show()
                            } else {
                                binding.tvStatus.text = "❌ $message"
                                Toast.makeText(this@MainActivity, "Error: $message", Toast.LENGTH_LONG).show()
                            }
                            addLog(message)
                        }
                    }
                }
            )
        }
    }

    private fun getFileName(uri: Uri): String {
        var fileName = "Unknown"
        contentResolver.query(uri, null, null, null, null)?.use { cursor ->
            val nameIndex = cursor.getColumnIndex(android.provider.OpenableColumns.DISPLAY_NAME)
            if (cursor.moveToFirst() && nameIndex >= 0) {
                fileName = cursor.getString(nameIndex)
            }
        }
        return fileName
    }

    private fun addLog(message: String) {
        val timestamp = SimpleDateFormat("HH:mm:ss", Locale.getDefault()).format(Date())
        val logEntry = "[$timestamp] $message\n"
        binding.tvLog.append(logEntry)

        // Auto-scroll to bottom
        binding.tvLog.post {
            val scrollView = binding.tvLog.parent as? android.widget.ScrollView
            scrollView?.fullScroll(View.FOCUS_DOWN)
        }
    }
}