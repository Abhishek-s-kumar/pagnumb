package com.example.pptxprocessor

import android.content.Context
import android.net.Uri
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.util.zip.ZipEntry
import java.util.zip.ZipInputStream
import java.util.zip.ZipOutputStream

class PPTXProcessor(private val context: Context) {

    interface ProcessingListener {
        fun onProgress(progress: Int, message: String)
        fun onComplete(success: Boolean, message: String)
    }

    suspend fun processPresentation(
        inputUri: Uri,
        outputUri: Uri,
        listener: ProcessingListener
    ) = withContext(Dispatchers.IO) {
        try {
            listener.onProgress(10, "Opening presentation...")

            // Create temporary files
            val tempDir = File(context.cacheDir, "pptx_temp_${System.currentTimeMillis()}")
            tempDir.mkdirs()

            val inputFile = File(tempDir, "input.pptx")
            val outputFile = File(tempDir, "output.pptx")

            // Copy input file to temp location
            context.contentResolver.openInputStream(inputUri)?.use { inputStream ->
                FileOutputStream(inputFile).use { outputStream ->
                    inputStream.copyTo(outputStream)
                }
            }

            listener.onProgress(20, "Extracting presentation...")

            // Extract PPTX (it's a ZIP file)
            val extractDir = File(tempDir, "extracted")
            extractDir.mkdirs()

            extractZip(inputFile, extractDir)

            listener.onProgress(40, "Processing slides...")

            // Find slide files
            val slideDir = File(extractDir, "ppt/slides")
            val slideFiles = slideDir.listFiles { file ->
                file.name.startsWith("slide") && file.name.endsWith(".xml")
            }?.sortedBy { it.name } ?: emptyList()

            listener.onProgress(50, "Adding slide numbers...")

            // Process each slide
            slideFiles.forEachIndexed { index, slideFile ->
                val slideNumber = index + 1
                addSlideNumber(slideFile, slideNumber)

                val progress = 50 + (index * 30 / slideFiles.size)
                listener.onProgress(progress, "Processing slide $slideNumber...")
            }

            listener.onProgress(80, "Rebuilding presentation...")

            // Repackage as PPTX
            createZip(extractDir, outputFile)

            listener.onProgress(90, "Saving file...")

            // Copy to output URI
            context.contentResolver.openOutputStream(outputUri)?.use { outputStream ->
                FileInputStream(outputFile).use { inputStream ->
                    inputStream.copyTo(outputStream)
                }
            }

            // Cleanup
            tempDir.deleteRecursively()

            listener.onProgress(100, "Processing complete!")
            listener.onComplete(true, "Successfully added slide numbers to ${slideFiles.size} slides")

        } catch (e: Exception) {
            e.printStackTrace()
            listener.onComplete(false, "Error: ${e.message}")
        }
    }

    private fun extractZip(zipFile: File, destDir: File) {
        ZipInputStream(FileInputStream(zipFile)).use { zipIn ->
            var entry: ZipEntry? = zipIn.nextEntry
            while (entry != null) {
                val file = File(destDir, entry.name)
                if (entry.isDirectory) {
                    file.mkdirs()
                } else {
                    file.parentFile?.mkdirs()
                    FileOutputStream(file).use { output ->
                        zipIn.copyTo(output)
                    }
                }
                zipIn.closeEntry()
                entry = zipIn.nextEntry
            }
        }
    }

    private fun createZip(sourceDir: File, zipFile: File) {
        ZipOutputStream(FileOutputStream(zipFile)).use { zipOut ->
            addDirToZip(sourceDir, sourceDir, zipOut)
        }
    }

    private fun addDirToZip(rootDir: File, sourceDir: File, zipOut: ZipOutputStream) {
        sourceDir.listFiles()?.forEach { file ->
            if (file.isDirectory) {
                addDirToZip(rootDir, file, zipOut)
            } else {
                val relativePath = rootDir.toURI().relativize(file.toURI()).path
                val entry = ZipEntry(relativePath)
                zipOut.putNextEntry(entry)
                FileInputStream(file).use { input ->
                    input.copyTo(zipOut)
                }
                zipOut.closeEntry()
            }
        }
    }

    private fun addSlideNumber(slideFile: File, slideNumber: Int) {
        try {
            // Read the slide XML
            val content = slideFile.readText()

            // Simple approach: Add text box with slide number
            // This is a simplified version - you might need to adjust based on your needs
            val slideNumberXml = """
                <p:sp>
                    <p:nvSpPr>
                        <p:cNvPr id="100$slideNumber" name="SlideNumber$slideNumber"/>
                        <p:cNvSpPr/>
                        <p:nvPr/>
                    </p:nvSpPr>
                    <p:spPr>
                        <a:xfrm>
                            <a:off x="8128000" y="6096000"/>
                            <a:ext cx="914400" cy="365760"/>
                        </a:xfrm>
                        <a:prstGeom prst="rect">
                            <a:avLst/>
                        </a:prstGeom>
                    </p:spPr>
                    <p:txBody>
                        <a:bodyPr wrap="none" rtlCol="0"/>
                        <a:lstStyle/>
                        <a:p>
                            <a:pPr algn="r"/>
                            <a:r>
                                <a:rPr lang="en-US" sz="1200" b="1">
                                    <a:solidFill>
                                        <a:srgbClr val="000000"/>
                                    </a:solidFill>
                                    <a:latin typeface="Arial"/>
                                </a:rPr>
                                <a:t>$slideNumber</a:t>
                            </a:r>
                        </a:p>
                    </p:txBody>
                </p:sp>
            """.trimIndent()

            // Find the closing tag of spTree and insert before it
            val modifiedContent = if (content.contains("</p:spTree>")) {
                content.replace("</p:spTree>", "$slideNumberXml</p:spTree>")
            } else {
                // Fallback: add before closing cSld tag
                content.replace("</p:cSld>", "$slideNumberXml</p:cSld>")
            }

            // Write back to file
            slideFile.writeText(modifiedContent)

        } catch (e: Exception) {
            // If XML manipulation fails, continue with other slides
            e.printStackTrace()
        }
    }
}