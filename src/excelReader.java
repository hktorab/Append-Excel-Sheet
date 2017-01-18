import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import java.io.File;
import java.io.IOException;

/**
 * Created by Torab on 19-Jan-17.
 */
public class excelReader {
    public static void main(String[] args) {
        File file=new File("firstexcel.xls");
        if(!file.exists())
        {
            try {
                create();
                excelWritingwriting();
            } catch (IOException e) {
                e.printStackTrace();
            } catch (WriteException e) {
                e.printStackTrace();
            } catch (BiffException e) {
                e.printStackTrace();
            }
        }
        try {
            excelWritingwriting();
        } catch (WriteException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        }

    }

    private static void excelWritingwriting() throws WriteException, IOException, BiffException {
        Workbook aWorkBook = Workbook.getWorkbook(new File("firstexcel.xls"));
        WritableWorkbook aCopy = Workbook.createWorkbook(new File("firstexcel.xls"), aWorkBook);

        WritableSheet aCopySheet = aCopy.getSheet(0);

        Label anotherWritableCell =  new Label(0,3,"Torab");



        aCopySheet.addCell(anotherWritableCell);

       // aCopySheet.addCell(anotherWritableCel2);

        aCopy.write();
        aCopy.close();

    }

    private static void create() throws IOException, WriteException {

        WritableWorkbook writableWorkbook = Workbook.createWorkbook(new File("firstexcel.xls"));
        writableWorkbook.createSheet("firstexcel.xls",0);

        writableWorkbook.write();
        writableWorkbook.close();
    }
}
