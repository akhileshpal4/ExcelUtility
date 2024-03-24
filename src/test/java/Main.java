import com.autoworld.excelreader.ReadExcel;

import java.io.IOException;

public class Main {
    public static void main(String[] args) throws IOException {
        System.out.println(ReadExcel.readData("xyz","sheet1"));
    }
}