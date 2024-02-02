package com.example.akhiltestapp

import android.os.Bundle
import android.widget.Button
import android.widget.EditText
import android.widget.TextView
import androidx.appcompat.app.AppCompatActivity
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import java.io.FileOutputStream
import java.io.IOException
import java.io.OutputStream
import java.util.Locale

class MainActivity : AppCompatActivity() {

    private lateinit var etMaterialType: EditText
    private lateinit var etCustomerName: EditText
    private lateinit var etWeight: EditText
    private lateinit var etImpurities: EditText
    private lateinit var tvNetWeight: TextView
    private lateinit var tvDalValue: TextView
    private lateinit var tvChuriValue: TextView
    private lateinit var tvGhatValue: TextView
    private lateinit var tvProcessingFees: TextView

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        etMaterialType = findViewById(R.id.etMaterialType)
        etCustomerName = findViewById(R.id.etCustomerName)
        etWeight = findViewById(R.id.etWeight)
        etImpurities = findViewById(R.id.etImpurities)
        tvNetWeight = findViewById(R.id.tvNetWeight)
        tvDalValue = findViewById(R.id.tvDalValue)
        tvChuriValue = findViewById(R.id.tvChuriValue)
        tvGhatValue = findViewById(R.id.tvGhatValue)
        tvProcessingFees = findViewById(R.id.tvProcessingFees)

        val calculateButton: Button = findViewById(R.id.calculateButton)
        calculateButton.setOnClickListener {
            calculateAndDisplay()
        }
    }

    private fun calculateAndDisplay() {
        // Get inputs
        val materialType: String = etMaterialType.text.toString()
        val customerName: String = etCustomerName.text.toString()
        val weight: Double = etWeight.text.toString().toDouble()
        val impurities: Double = etImpurities.text.toString().toDouble()

        // Calculate net weight
        val netWeight = weight - impurities

        // Calculate values based on material type
        var dalValue = 0.0
        var churiValue = 0.0
        var ghatValue = 0.0
        var processingFees = 0.0
        if (materialType.equals("toor", ignoreCase = true)) {
            dalValue = netWeight * 0.65
            churiValue = netWeight * 0.32
            ghatValue = netWeight * 0.03
            processingFees = netWeight * 6
        } else if (materialType.equals("chana", ignoreCase = true)) {
            // Adjust multiplication factors for chana if needed
            dalValue = netWeight * 0.6
            churiValue = netWeight * 0.35
            ghatValue = netWeight * 0.05
            processingFees = netWeight * 6 // Adjust as needed for chana
        }

        // Display results
        tvNetWeight.text = String.format(Locale.getDefault(), "Net Weight: %.2f", netWeight)
        tvDalValue.text = String.format(Locale.getDefault(), "Dal Value: %.2f", dalValue)
        tvChuriValue.text = String.format(Locale.getDefault(), "Churi Value: %.2f", churiValue)
        tvGhatValue.text = String.format(Locale.getDefault(), "Ghat Value: %.2f", ghatValue)
        tvProcessingFees.text =
            String.format(Locale.getDefault(), "Processing Fees: %.2f", processingFees)

        // Save results to Excel sheet
        saveToExcel(
            customerName,
            materialType,
            netWeight,
            dalValue,
            churiValue,
            ghatValue,
            processingFees
        )
    }

    private fun saveToExcel(
        customerName: String,
        materialType: String,
        netWeight: Double,
        dalValue: Double,
        churiValue: Double,
        ghatValue: Double,
        processingFees: Double
    ) {
        // Use Apache POI library to create an Excel sheet and save the results

        // Create a new workbook and sheet
        val workbook: Workbook = HSSFWorkbook()
        val sheet: Sheet = workbook.createSheet("Results")

        // Create header row
        val headerRow: Row = sheet.createRow(0)
        headerRow.createCell(0).setCellValue("Customer Name")
        headerRow.createCell(1).setCellValue("Material Type")
        headerRow.createCell(2).setCellValue("Net Weight")
        headerRow.createCell(3).setCellValue("Dal Value")
        headerRow.createCell(4).setCellValue("Churi Value")
        headerRow.createCell(5).setCellValue("Ghat Value")
        headerRow.createCell(6).setCellValue("Processing Fees")

        // Create data row
        val dataRow: Row = sheet.createRow(1)
        dataRow.createCell(0).setCellValue(customerName)
        dataRow.createCell(1).setCellValue(materialType)
        dataRow.createCell(2).setCellValue(netWeight)
        dataRow.createCell(3).setCellValue(dalValue)
        dataRow.createCell(4).setCellValue(churiValue)
        dataRow.createCell(5).setCellValue(ghatValue)
        dataRow.createCell(6).setCellValue(processingFees)

        // Save the workbook to a file
        try {
            val filePath: String = getExternalFilesDir(null).toString() + "/results.xls"
            val outputStream: OutputStream = FileOutputStream(filePath)
            workbook.write(outputStream)
            outputStream.close()
        } catch (e: IOException) {
            e.printStackTrace()
        }
    }
}