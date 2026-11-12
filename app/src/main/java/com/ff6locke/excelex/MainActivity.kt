package com.ff6locke.excelex

import android.app.Activity
import android.content.Intent
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
    private var rows = mutableListOf(mutableListOf(""))
    private var fileName = "Workbook.xlsx"
    private var pendingSave = false

    override fun onCreate(savedInstanceState: Bundle?) { super.onCreate(savedInstanceState); home(); handleIntent(intent) }
    override fun onNewIntent(intent: Intent) { super.onNewIntent(intent); setIntent(intent); handleIntent(intent) }

    private fun home() {
        val root = LinearLayout(this).apply { orientation = LinearLayout.VERTICAL; setBackgroundColor(0xFFF5F7F5.toInt()) }
        val bar = LinearLayout(this).apply { gravity = Gravity.CENTER_VERTICAL; setPadding(12, 10, 8, 10); setBackgroundColor(0xFF087C3E.toInt()) }
        bar.addView(TextView(this).apply { text = "ExcelEX"; textSize = 24f; setTextColor(-1); setTypeface(null, 1) }, LinearLayout.LayoutParams(0, -2, 1f))
        bar.addView(Button(this).apply { text = "Open"; setOnClickListener { chooseFile() } })
        bar.addView(Button(this).apply { text = "Save"; setOnClickListener { chooseSave() } })
        root.addView(bar)
        status = TextView(this).apply { text = "Open an Excel workbook or tap a cell to start"; setPadding(12, 10, 12, 10); textSize = 13f }
        root.addView(status)
        sheet = TableLayout(this)
        val horizontal = HorizontalScrollView(this)
        val vertical = ScrollView(this); vertical.addView(sheet); horizontal.addView(vertical)
        root.addView(horizontal, LinearLayout.LayoutParams(-1, 0, 1f))
        root.addView(Button(this).apply { text = "+ Add row"; setOnClickListener { rows.add(MutableList((rows.firstOrNull()?.size ?: 4).coerceAtLeast(1)) { "" }); render() } })
        setContentView(root); render()
    }

    private fun chooseFile() {
        val i = Intent(Intent.ACTION_OPEN_DOCUMENT).apply {
            addCategory(Intent.CATEGORY_OPENABLE); type = "*/*"
            putExtra(Intent.EXTRA_MIME_TYPES, arrayOf("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "application/vnd.ms-excel", "text/csv"))
        }
        startActivityForResult(i, 100)
    }
    private fun chooseSave() {
        val i = Intent(Intent.ACTION_CREATE_DOCUMENT).apply {
            addCategory(Intent.CATEGORY_OPENABLE); type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            putExtra(Intent.EXTRA_TITLE, if (fileName.endsWith(".xlsx", true)) fileName else "ExcelEX_Workbook.xlsx")
        }
        startActivityForResult(i, 101)
    }
    @Deprecated("Legacy activity result callback")
    override fun onActivityResult(requestCode: Int, resultCode: Int, data: Intent?) {
        super.onActivityResult(requestCode, resultCode, data)
        if (resultCode == Activity.RESULT_OK) data?.data?.let { if (requestCode == 100) openFile(it) else if (requestCode == 101) writeFile(it) }
    }
    private fun handleIntent(i: Intent?) { if (i?.action == Intent.ACTION_VIEW) i.data?.let { openFile(it) } }

    private fun openFile(uri: Uri) {
        try {
            fileName = contentResolver.query(uri, null, null, null, null)?.use { c ->
                val ix = c.getColumnIndex(android.provider.OpenableColumns.DISPLAY_NAME); if (ix >= 0 && c.moveToFirst()) c.getString(ix) else "Workbook"
            } ?: "Workbook"
            val ext = fileName.substringAfterLast('.', "").lowercase(Locale.ROOT)
            contentResolver.openInputStream(uri).use { stream ->
                requireNotNull(stream)
                when (ext) {
                    "csv" -> rows = BufferedReader(InputStreamReader(stream, Charsets.UTF_8)).readLines().map { parseCsv(it).toMutableList() }.toMutableList()
                    "xls", "xlsx" -> {
                        val wb = WorkbookFactory.create(stream); val s = wb.getSheetAt(0); rows = mutableListOf()
                        val count = (0..s.lastRowNum).maxOfOrNull { s.getRow(it)?.lastCellNum?.toInt() ?: 0 } ?: 1
                        for (r in 0..s.lastRowNum) rows.add(MutableList(count.coerceAtLeast(1)) { c -> s.getRow(r)?.getCell(c)?.toString() ?: "" })
                        wb.close()
                    }
                    else -> error("Supported formats: .xlsx, .xls and .csv")
                }
            }
            if (rows.isEmpty()) rows.add(mutableListOf(""))
            status.text = "$fileName • ${rows.size} rows"
            render()
        } catch (e: Exception) { Toast.makeText(this, "Open failed: ${e.message}", Toast.LENGTH_LONG).show() }
    }
    private fun parseCsv(line: String): List<String> = line.split(Regex(",(?=(?:[^"]*"[^"]*")*[^"]*$)")).map { it.trim().removeSurrounding("\"").replace("\"\"", "\"") }
    private fun render() {
        sheet.removeAllViews()
        val cols = (rows.maxOfOrNull { it.size } ?: 1).coerceIn(1, 20)
        rows.forEachIndexed { r, values ->
            val tr = TableRow(this)
            repeat(cols) { c ->
                val cell = TextView(this).apply {
                    text = values.getOrNull(c).orEmpty(); textSize = 14f; gravity = Gravity.CENTER_VERTICAL
                    setPadding(12, 14, 12, 14); minWidth = 110
                    setTextColor(0xFF17251C.toInt()); setBackgroundColor(if (r == 0) 0xFFE2F2E8.toInt() else if (r % 2 == 0) -1 else 0xFFF8FBF9.toInt())
                    setOnClickListener { editCell(r, c) }
                }
                tr.addView(cell, TableRow.LayoutParams(110, -2))
            }
            sheet.addView(tr)
        }
    }
    private fun editCell(r: Int, c: Int) {
        while (rows.size <= r) rows.add(mutableListOf())
        while (rows[r].size <= c) rows[r].add("")
        val input = EditText(this).apply { setText(rows[r][c]); hint = "Cell value" }
        android.app.AlertDialog.Builder(this).setTitle("Edit cell").setView(input).setNegativeButton("Cancel", null)
            .setPositiveButton("Apply") { _, _ -> rows[r][c] = input.text.toString(); render() }.show()
    }
    private fun writeFile(uri: Uri) {
        try {
            val wb = XSSFWorkbook(); val s = wb.createSheet("Sheet1")
            rows.forEachIndexed { r, values -> val row = s.createRow(r); values.forEachIndexed { c, v -> row.createCell(c).setCellValue(v) } }
            contentResolver.openOutputStream(uri).use { out -> requireNotNull(out); wb.write(out) }; wb.close()
            Toast.makeText(this, "Workbook saved", Toast.LENGTH_SHORT).show()
        } catch (e: Exception) { Toast.makeText(this, "Save failed: ${e.message}", Toast.LENGTH_LONG).show() }
    }
}
