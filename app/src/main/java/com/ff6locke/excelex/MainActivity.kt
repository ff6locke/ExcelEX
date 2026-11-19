package com.ff6locke.excelex

import android.app.Activity
import android.content.ClipData
import android.content.ClipboardManager
import android.content.Context
import android.content.Intent
import android.graphics.Color
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
    private lateinit var verticalScroll: ScrollView
    private var rows = MutableList(20) { MutableList(8) { "" } }
    private var fileName = "Workbook.xlsx"
    private var selectedCell: Pair<Int, Int>? = null
    private var searchMatches = mutableListOf<Pair<Int, Int>>()
    private var searchIndex = -1
    private var highlightedCell: TextView? = null

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
            text = "Open a workbook, then tap a cell. Long-press a cell to select it for copy/paste."
            setPadding(12, 10, 12, 10)
            textSize = 13f
        }
        root.addView(status)

        val searchBar = LinearLayout(this).apply {
            orientation = LinearLayout.HORIZONTAL
            gravity = Gravity.CENTER_VERTICAL
            setPadding(8, 2, 8, 6)
        }
        val searchInput = EditText(this).apply {
            hint = "Search in workbook"
            singleLine = true
            textSize = 15f
            setPadding(10, 6, 10, 6)
        }
        searchBar.addView(searchInput, LinearLayout.LayoutParams(0, dp(48), 1f))
        searchBar.addView(Button(this).apply {
            text = "Find"
            setOnClickListener {
                searchMatches.clear()
                searchIndex = -1
                findNext(searchInput.text.toString())
            }
        })
        searchBar.addView(Button(this).apply {
            text = "Next"
            setOnClickListener { findNext(searchInput.text.toString()) }
        })
        root.addView(searchBar)

        sheet = TableLayout(this)
        val horizontal = HorizontalScrollView(this)
        verticalScroll = ScrollView(this)
        verticalScroll.addView(sheet)
        horizontal.addView(verticalScroll)
        root.addView(horizontal, LinearLayout.LayoutParams(-1, 0, 1f))

        val actions = LinearLayout(this).apply { orientation = LinearLayout.HORIZONTAL }
        actions.addView(Button(this).apply {
            text = "Copy"
            setOnClickListener { copySelectedCell() }
        }, LinearLayout.LayoutParams(0, -2, 1f))
        actions.addView(Button(this).apply {
            text = "Paste"
            setOnClickListener { pasteIntoSelectedCell() }
        }, LinearLayout.LayoutParams(0, -2, 1f))
        actions.addView(Button(this).apply {
            text = "+ Row"
            setOnClickListener {
                val columnCount = (rows.maxOfOrNull { it.size } ?: 8).coerceAtLeast(1)
                rows.add(MutableList(columnCount) { "" })
                render()
                status.text = rows.size.toString() + " rows • New row added"
            }
        }, LinearLayout.LayoutParams(0, -2, 1f))
        actions.addView(Button(this).apply {
            text = "+ Col"
            setOnClickListener {
                rows.forEach { it.add("") }
                render()
                status.text = rows.size.toString() + " rows • " + (rows.maxOfOrNull { it.size } ?: 0) + " columns"
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
            selectedCell = null
            searchMatches.clear()
            status.text = fileName + " • " + (rows.size - 20) + " data rows • 20 blank rows ready"
            render()
        } catch (e: Exception) {
            Toast.makeText(this, "Open failed: " + e.message, Toast.LENGTH_LONG).show()
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
        highlightedCell = null
        val columnCount = (rows.maxOfOrNull { it.size } ?: 1).coerceIn(1, 20)
        rows.forEachIndexed { rowIndex, values ->
            val tableRow = TableRow(this)
            repeat(columnCount) { columnIndex ->
                val cell = TextView(this).apply {
                    text = values.getOrNull(columnIndex).orEmpty()
                    textSize = 14f
                    gravity = Gravity.CENTER_VERTICAL
                    setPadding(12, 14, 12, 14)
                    minWidth = dp(110)
                    setTextColor(0xFF17251C.toInt())
                    background = cellBackground(rowIndex, false)
                    setOnClickListener {
                        selectedCell = rowIndex to columnIndex
                        editCell(rowIndex, columnIndex)
                    }
                    setOnLongClickListener {
                        selectedCell = rowIndex to columnIndex
                        status.text = "Selected " + cellAddress(rowIndex, columnIndex) + " • Use Copy or Paste"
                        Toast.makeText(this@MainActivity, "Cell " + cellAddress(rowIndex, columnIndex) + " selected", Toast.LENGTH_SHORT).show()
                        true
                    }
                }
                tableRow.addView(cell, TableRow.LayoutParams(dp(110), -2))
            }
            sheet.addView(tableRow)
        }
    }

    private fun cellBackground(rowIndex: Int, highlighted: Boolean): GradientDrawable =
        GradientDrawable().apply {
            setColor(
                if (highlighted) 0xFFFFF2A8.toInt()
                else if (rowIndex == 0) 0xFFE2F2E8.toInt()
                else if (rowIndex % 2 == 0) Color.WHITE else 0xFFF8FBF9.toInt()
            )
            setStroke(dp(if (highlighted) 2 else 1), if (highlighted) 0xFF168447.toInt() else 0xFFD2DAD5.toInt())
        }

    private fun findNext(query: String) {
        val term = query.trim()
        if (term.isEmpty()) {
            Toast.makeText(this, "Enter text to search", Toast.LENGTH_SHORT).show()
            return
        }
        if (searchMatches.isEmpty()) {
            rows.forEachIndexed { r, row ->
                row.forEachIndexed { c, value ->
                    if (value.contains(term, ignoreCase = true)) searchMatches.add(r to c)
                }
            }
        }
        if (searchMatches.isEmpty()) {
            status.text = "No matches for \"" + term + "\""
            Toast.makeText(this, "No matches found", Toast.LENGTH_SHORT).show()
            return
        }
        searchIndex = (searchIndex + 1) % searchMatches.size
        val match = searchMatches[searchIndex]
        selectedCell = match
        val cell = sheet.getChildAt(match.first)?.let { it as? TableRow }?.getChildAt(match.second) as? TextView
        cell?.let {
            highlightedCell?.let { previous ->
                val previousRow = previous.parent as? TableRow
                val rowIndex = if (previousRow != null) sheet.indexOfChild(previousRow) else 0
                previous.background = cellBackground(rowIndex.coerceAtLeast(0), false)
            }
            it.background = cellBackground(match.first, true)
            highlightedCell = it
            it.post {
                val row = it.parent as? TableRow
                verticalScroll.smoothScrollTo(0, row?.top ?: it.top)
            }
        }
        status.text = "Found " + (searchIndex + 1) + " of " + searchMatches.size + " • Cell " + cellAddress(match.first, match.second)
    }

    private fun copySelectedCell() {
        val position = selectedCell
        if (position == null) {
            Toast.makeText(this, "Tap or long-press a cell first", Toast.LENGTH_SHORT).show()
            return
        }
        val value = rows.getOrNull(position.first)?.getOrNull(position.second).orEmpty()
        val clipboard = getSystemService(Context.CLIPBOARD_SERVICE) as ClipboardManager
        clipboard.setPrimaryClip(ClipData.newPlainText("ExcelEX cell", value))
        Toast.makeText(this, "Copied " + cellAddress(position.first, position.second), Toast.LENGTH_SHORT).show()
    }

    private fun pasteIntoSelectedCell() {
        val position = selectedCell
        if (position == null) {
            Toast.makeText(this, "Tap or long-press a cell first", Toast.LENGTH_SHORT).show()
            return
        }
        val clipboard = getSystemService(Context.CLIPBOARD_SERVICE) as ClipboardManager
        val clip = clipboard.primaryClip
        if (clip == null || clip.itemCount == 0) {
            Toast.makeText(this, "Clipboard is empty", Toast.LENGTH_SHORT).show()
            return
        }
        val pastedText = clip.getItemAt(0).coerceToText(this)?.toString() ?: ""
        val pasteRows = pastedText.replace("\r\n", "\n").replace('\r', '\n').split('\n')
        pasteRows.forEachIndexed { rowOffset, line ->
            val values = line.split('\t')
            val targetRow = position.first + rowOffset
            while (rows.size <= targetRow) rows.add(MutableList((rows.maxOfOrNull { it.size } ?: 8).coerceAtLeast(8)) { "" })
            values.forEachIndexed { columnOffset, value ->
                val targetColumn = position.second + columnOffset
                while (rows[targetRow].size <= targetColumn) rows[targetRow].add("")
                rows[targetRow][targetColumn] = value
            }
        }
        render()
        status.text = "Pasted into " + cellAddress(position.first, position.second)
    }

    private fun cellAddress(row: Int, column: Int): String {
        var number = column + 1
        var letters = ""
        while (number > 0) {
            val remainder = (number - 1) % 26
            letters = ('A' + remainder) + letters
            number = (number - 1) / 26
        }
        return letters + (row + 1)
    }

    private fun dp(value: Int): Int = (value * resources.displayMetrics.density).toInt().coerceAtLeast(1)

    private fun editCell(rowIndex: Int, columnIndex: Int) {
        while (rows.size <= rowIndex) rows.add(mutableListOf())
        while (rows[rowIndex].size <= columnIndex) rows[rowIndex].add("")
        val input = EditText(this).apply {
            setText(rows[rowIndex][columnIndex])
            hint = "Cell value"
            setSingleLine(false)
        }
        android.app.AlertDialog.Builder(this)
            .setTitle("Edit " + cellAddress(rowIndex, columnIndex))
            .setView(input)
            .setNegativeButton("Cancel", null)
            .setNeutralButton("Paste") { _, _ ->
                val clipboard = getSystemService(Context.CLIPBOARD_SERVICE) as ClipboardManager
                val clip = clipboard.primaryClip
                if (clip != null && clip.itemCount > 0) {
                    rows[rowIndex][columnIndex] = clip.getItemAt(0).coerceToText(this)?.toString().orEmpty()
                    render()
                }
            }
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
            Toast.makeText(this, "Save failed: " + e.message, Toast.LENGTH_LONG).show()
        }
    }
}
