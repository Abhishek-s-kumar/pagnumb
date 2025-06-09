package com.example.pptxprocessor

import android.view.LayoutInflater
import android.view.View
import android.widget.Button
import android.widget.ProgressBar
import android.widget.TextView

// Simple ViewBinding implementation
class ActivityMainBinding private constructor(
    private val rootView: View
) {
    val btnSelectFile: Button = rootView.findViewById(R.id.btnSelectFile)
    val tvSelectedFile: TextView = rootView.findViewById(R.id.tvSelectedFile)
    val btnProcessFile: Button = rootView.findViewById(R.id.btnProcessFile)
    val progressBar: ProgressBar = rootView.findViewById(R.id.progressBar)
    val tvStatus: TextView = rootView.findViewById(R.id.tvStatus)
    val tvLog: TextView = rootView.findViewById(R.id.tvLog)

    fun getRoot(): View = rootView

    companion object {
        fun inflate(inflater: LayoutInflater): ActivityMainBinding {
            val view = inflater.inflate(R.layout.activity_main, null)
            return ActivityMainBinding(view)
        }
    }
}