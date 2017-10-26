package com.example.jade.readexceldemo;

import android.content.res.AssetManager;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;


public class MainActivity extends AppCompatActivity {

    private Button btnRead;
    private AssetManager manager;

    @Override
    protected void onCreate(Bundle savedInstanceState)
    {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        btnRead = (Button)findViewById(R.id.btnRead);

        // An asset manager that will be used to grab the excel file from the assets folder
        manager = getAssets();

        btnRead.setOnClickListener(new View.OnClickListener()
        {
            @Override
            public void onClick(View v)
            {
                readExcel("Board-Games.xls");
            }
        });
    }

    // This method will read data from an excel file and output the data to the logcat.
    // IMPORTANT: Excel files MUST be saved as Excel 97-2003 Workbook (*.xls) in order to be read.
    //
    // filename: The name of the excel file to be read.
    public void readExcel(String filename)
    {
        try
        {
            // Open the file from the assets folder
            InputStream stream = manager.open(filename);

            // Create a file system object
            POIFSFileSystem fileSystem = new POIFSFileSystem(stream);

            // Create a workbook using the file system
            HSSFWorkbook workBook = new HSSFWorkbook(fileSystem);

            // Get the first sheet from the workbook
            HSSFSheet workSheet = workBook.getSheetAt(13);

            // Iterator to iterate over each row in the worksheet
            Iterator rowIterator = workSheet.rowIterator();

            // While the worksheet has rows with data...
            while(rowIterator.hasNext())
            {
                // Assign the current row to this row object variable
                HSSFRow row = (HSSFRow) rowIterator.next();

                // Iterator to iterate over each cell in the row
                Iterator cellIterator = row.cellIterator();

                // While the current row has a cell that contains data...
                while(cellIterator.hasNext())
                {
                    // Assign the current cell to this cell object variable
                    HSSFCell cell = (HSSFCell) cellIterator.next();

                    // Output the contents in the current cell to the logcat
                    Log.d("Excel Data:", "Cell Value: " + cell.toString());
                }
            }

        }
        catch(IOException e)
        {
            // Hey I just caught you
            // And this is crazy
            // But here's your stack trace
            // So try me maybe
            e.printStackTrace();
        }
    }
}
