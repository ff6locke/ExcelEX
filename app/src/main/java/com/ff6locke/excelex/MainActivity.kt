package com.ff6locke.excelex

import android.app.Activity
import android.content.Intent
import android.graphics.drawable.GradientDrawable
import android.net.Uri
import android.os.Bundle
import android.view.Gravity
import android.widget.*
import androidx.appcompat.app.AppCompatActivity
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.BufferedReader
import java.io.InputStreamReader
import java.util.Locale

class MainActivity : AppCompatActivity() {
    private lateinit var sheet: TableLayout
    private lateinit var status: TextView
    private var rows = MutableList(20) { MutableList(8) { "" } }
    private var fileName = "Workbook.xlsx"

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        home()
        handleIntent(intent)
    }

    override fun onNewIntent(intent: Intent) {
        super.onNewIntent(intent)
        setIntent(intent)
        handleIntent(intent)
    }

    private fun home() {
        val root = LinearLayout(this).apply {
            orientation = LinearLayout.VERTICAL
            setBackgroundColor(0xFFF5F7F5.toInt())
        }
        val bar = LinearLayout(this).apply {
            gravity = Gravity.CENTER_VERTICAL
            setPadding(12, 10, 8, 10)
            setBackgroundColor(0xFF087C3E.toInt())
        }
        bar.addView(TextView(this).apply {
            text = "ExcelEX"
            textSize = 24f
            setTextColor(-1)
            setTypeface(null, 1)
        }, LinearLayout.LayoutParams(0, -2, 1f))
        bar.addView(Button(this).apply {
            text = "Open"
            setOnClickListener { chooseFile() }
        })
        bar.addView(Button(this).apply {
            text = "Save"
            setOnClickListener { chooseSave() }
        })
        root.addView(bar)

        status = TextView(this).apply {
            text = "Open a workbook, then tap any cell. 20 blank rows are ready to edit."
            setPadding(12, 10, 12, 10)
            textSize = 13f
        }
        root.addView(status)

        sheet = TableLayout(this)
        val horizontal = HorizontalScrollView(this)
        val vertical = ScrollView(this)
        vertical.addView(sheet)
        horizontal.addView(vertical)
        root.addView(horizontal, LinearLayout.LayoutParams(-1, 0, 1f))
        val actions = LinearLayout(this).apply { orientation = LinearLayout.HORIZONTAL }
        actions.addView(Button(this).apply {
            text = "+ Add row"
            setOnClickListener {
                val columnCount = (rows.maxOfOrNull { it.size } ?: 8).coerceAtLeast(1)
                rows.add(MutableList(columnCount) { "" })
                render()
                status.text = "${rows.size} rows • New row added"
            }
        }, LinearLayout.LayoutParams(0, -2, 1f))
        actions.addView(Button(this).apply {
            text = "+ Add column"
            setOnClickListener {
                rows.forEach { it.add("") }
                render()
                status.text = "${rows.size} rows • ${rows.maxOfOrNull { it.size } ?: 0} columns"
            }
        }, LinearLayout.LayoutParams(0, -2, 1f))
        root.addView(actions)
        setContentView(root)
        render()
    }

    private fun chooseFile() {
        val i = Intent(Intent.ACTION_OPEN_DOCUMENT).apply {
            addCategory(Intent.CATEGORY_OPENABLE)
            type = "*/*"
            putExtra(Intent.EXTRA_MIME_TYPES, arrayOf(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "application/vnd.ms-excel",
                "text/csv"
            ))
        }
        startActivityForResult(i, 100)
    }

    private fun chooseSave() {
        val i = Intent(Intent.ACTION_CREATE_DOCUMENT).apply {
            addCategory(Intent.CATEGORY_OPENABLE)
            type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            putExtra(Intent.EXTRA_TITLE, if (fileName.endsWith(".xlsx", true)) fileName else "ExcelEX_Workbook.xlsx")
        }
        startActivityForResult(i, 101)
    }

    @Deprecated("Legacy activity result callback")
    override fun onActivityResult(requestCode: Int, resultCode: Int, data: Intent?) {
        super.onActivityResult(requestCode, resultCode, data)
        if (resultCode == Activity.RESULT_OK) {
            data?.data?.let { uri ->
                when (requestCode) {
                    100 -> openFile(uri)
                    101 -> writeFile(uri)
                }
            }
        }
    }

    private fun handleIntent(i: Intent?) {
        if (i?.action == Intent.ACTION_VIEW) i.data?.let { openFile(it) }
    }

    private fun openFile(uri: Uri) {
        try {
            fileName = contentResolver.query(uri, null, null, null, null)?.use { cursor ->
                val index = cursor.getColumnIndex(android.provider.OpenableColumns.DISPLAY_NAME)
                if (index >= 0 && cursor.moveToFirst()) cursor.getString(index) else "Workbook"
            } ?: "Workbook"

            val extension = fileName.substringAfterLast('.', "").lowercase(Locale.ROOT)
            contentResolver.openInputStream(uri).use { stream ->
                requireNotNull(stream)
                when (extension) {
                    "csv" -> {
                        rows = BufferedReader(InputStreamReader(stream, Charsets.UTF_8))
                            .readLines()
                            .map { parseCsv(it).toMutableList() }
                            .toMutableList()
                    }
                    "xls", "xlsx" -> {
                        WorkbookFactory.create(stream).use { workbook ->
                            val firstSheet = workbook.getSheetAt(0)
                            rows = mutableListOf()
                            val columnCount = (0..firstSheet.lastRowNum)
                                .maxOfOrNull { firstSheet.getRow(it)?.lastCellNum?.toInt() ?: 0 }
                                ?: 1
                            for (rowIndex in 0..firstSheet.lastRowNum) {
                                val sourceRow = firstSheet.getRow(rowIndex)
                                rows.add(MutableList(columnCount.coerceAtLeast(1)) { columnIndex ->
                                    sourceRow?.getCell(columnIndex)?.toString() ?: ""
                                })
                            }
                        }
                    }
                    else -> error("Supported formats: .xlsx, .xls and .csv")
                }
            }
            if (rows.isEmpty()) rows.add(mutableListOf(""))
            val width = (rows.maxOfOrNull { it.size } ?: 1).coerceAtLeast(8)
            rows.forEach { row -> while (row.size < width) row.add("") }
            repeat(20) { rows.add(MutableList(width) { "" }) }
            status.text = "$fileName • ${rows.size - 20} data rows • 20 blank rows ready"
            render()
        } catch (e: Exception) {
            Toast.makeText(this, "Open failed: ${e.message}", Toast.LENGTH_LONG).show()
        }
    }

    private fun parseCsv(line: String): List<String> {
        val fields = mutableListOf<String>()
        val field = StringBuilder()
        var quoted = false
        var index = 0
        while (index < line.length) {
            val character = line[index]
            when {
                character == '"' && quoted && index + 1 < line.length && line[index + 1] == '"' -> {
                    field.append('"')
                    index++
                }
                character == '"' -> quoted = !quoted
                character == ',' && !quoted -> {
                    fields.add(field.toString())
                    field.setLength(0)
                }
                else -> field.append(character)
            }
            index++
        }
        fields.add(field.toString())
        return fields
    }

    private fun render() {
        sheet.removeAllViews()
        val columnCount = (rows.maxOfOrNull { it.size } ?: 1).coerceIn(1, 20)
        rows.forEachIndexed { rowIndex, values ->
            val tableRow = TableRow(this)
            repeat(columnCount) { columnIndex ->
                val cell = TextView(this).apply {
                    text = values.getOrNull(columnIndex).orEmpty()
                    textSize = 14f
                    gravity = Gravity.CENTER_VERTICAL
                    setPadding(12, 14, 12, 14)
                    minWidth = 110
                    setTextColor(0xFF17251C.toInt())
                    background = GradientDrawable().apply {
                        setColor(
                            if (rowIndex == 0) 0xFFE2F2E8.toInt()
                            else if (rowIndex % 2 == 0) -1 else 0xFFF8FBF9.toInt()
                        )
                        setStroke(dp(1), 0xFFD2DAD5.toInt())
                    }
                    setOnClickListener { editCell(rowIndex, columnIndex) }
                }
                tableRow.addView(cell, TableRow.LayoutParams(110, -2))
            }
            sheet.addView(tableRow)
        }
    }

    private fun dp(value: Int): Int = (value * resources.displayMetrics.density).toInt().coerceAtLeast(1)

    private fun editCell(rowIndex: Int, columnIndex: Int) {
        while (rows.size <= rowIndex) rows.add(mutableListOf())
        while (rows[rowIndex].size <= columnIndex) rows[rowIndex].add("")
        val input = EditText(this).apply {
            setText(rows[rowIndex][columnIndex])
            hint = "Cell value"
        }
        android.app.AlertDialog.Builder(this)
            .setTitle("Edit cell")
            .setView(input)
            .setNegativeButton("Cancel", null)
            .setPositiveButton("Apply") { _, _ ->
                rows[rowIndex][columnIndex] = input.text.toString()
                render()
            }
            .show()
    }

    private fun writeFile(uri: Uri) {
        try {
            XSSFWorkbook().use { workbook ->
                val outputSheet = workbook.createSheet("Sheet1")
                rows.forEachIndexed { rowIndex, values ->
                    val outputRow = outputSheet.createRow(rowIndex)
                    values.forEachIndexed { columnIndex, value ->
                        outputRow.createCell(columnIndex).setCellValue(value)
                    }
                }
                contentResolver.openOutputStream(uri).use { output ->
                    requireNotNull(output)
                    workbook.write(output)
                }
            }
            Toast.makeText(this, "Workbook saved", Toast.LENGTH_SHORT).show()
        } catch (e: Exception) {
            Toast.makeText(this, "Save failed: ${e.message}", Toast.LENGTH_LONG).show()
        }
    }
}
