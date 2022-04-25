import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

public class ExcelHelper {
    public HashMap<String, String> getSpareData(String spare) {
        HashMap<String, String> data = new HashMap<String, String>();
        String[] spareDetails = new String[5];
        String[] identifier = new String[4];
        try {
            String path = System.getProperty("user.dir") + "\\src\\main\\resources\\Data\\SparesList.xlsx";
            System.out.println("PAHT: " + path);
            File file = new File(path);   //creating a new file instance
            FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object
            Iterator<Row> itr = sheet.iterator();    //iterating over excel file
            int i = 0;
            while (itr.hasNext()) {
                Row row = itr.next();
                Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column
                if (row.getCell(0).toString().toLowerCase().equals(spare.toLowerCase())) {
                    int j = 0;
                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        switch (cell.getCellType()) {
                            case NUMERIC:
                                spareDetails[j] = String.valueOf(cell.getNumericCellValue());
                                break;
                            case STRING:
                                spareDetails[j] = String.valueOf(cell.getStringCellValue());
                                break;
                        }
                        j++;
                    }
                    for (int k = 0; k <= spareDetails.length; k++) {
                        if (k == 0)
                            data.put("SpareName", spareDetails[k]);
                        else if (k == 1)
                            data.put("Vendor", spareDetails[k]);
                        else if (k == 2)
                            data.put("Quantity", spareDetails[k]);
                        else if (k == 3)
                            data.put("Price", spareDetails[k]);
                        else if (k == 4)
                            data.put("Address", spareDetails[k]);
                    }
                }
            }
            fis.close();
            wb.close();
        } catch (Exception e) {
            System.out.println("Error occured reading excel file ");
        }

        return data;
    }

    public void addSparesData(String[] inputData) throws IOException {
        try {
            FileInputStream inputStream = new FileInputStream(new File(System.getProperty("user.dir") + "\\src\\main\\resources\\Data\\SparesList.xlsx"));
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheetAt(0);
            int rowCount = sheet.getLastRowNum();
            System.out.println("last row in data: " + rowCount);
            Row row = sheet.createRow(++rowCount);

            for (int i = 0; i < inputData.length; i++) {
                Cell cell = row.createCell(i);
                cell.setCellValue(inputData[i]);
            }
            inputStream.close();
            FileOutputStream outputStream = new FileOutputStream(System.getProperty("user.dir") + "\\src\\main\\resources\\Data\\SparesList.xlsx");
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();
        } catch (IOException | EncryptedDocumentException ex) {
            assert ex != null;
            ex.printStackTrace();
            System.out.println("Error Occured while adding data");
        }
    }

    public HashMap<String, String> getSpareDataForSales(String spareName) throws IOException {
        HashMap<String, String> data = new HashMap<String, String>();
        String amount = "";
        int remainingQuantity = 0;
        Scanner sc = new Scanner(System.in);
        data = getSpareData(spareName);
        System.out.println("Price for each unit is: " + data.get("Price"));
        System.out.println("Please enter quantity you want to buy below this number: " + data.get("Quantity"));
        Integer quantity = sc.nextInt();
        if (quantity <= Integer.parseInt(data.get("Quantity").split("\\.")[0]) && quantity >= 0) {
            amount = String.valueOf(quantity * Integer.parseInt(data.get("Price").split("\\.")[0]));
            System.out.println("Please pay the following amount: " + amount);
            remainingQuantity = Integer.parseInt(data.get("Quantity").split("\\.")[0]) - quantity;
            System.out.println("remaining Quantity: " + remainingQuantity);
            addSalesSparesData(spareName, quantity, Integer.parseInt(amount));
            updateQuantity(spareName, remainingQuantity);
        } else if (quantity == 0)
            System.out.println("Cannot process your request please try again");
        else
            System.out.println("Cannot process your request please try again");
        return data;
    }

    public void addSalesSparesData(String spareName, int quantity, int price) throws IOException {
        try {
            FileInputStream inputStream = new FileInputStream(new File(System.getProperty("user.dir") + "\\src\\main\\resources\\Data\\daysSale.xlsx"));
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheetAt(0);
            int rowCount = sheet.getLastRowNum();
            Date date = new Date();
            SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
            String saleDate = formatter.format(date);
            System.out.println("last row in data: " + rowCount);
            Row row = sheet.createRow(++rowCount);
            Cell cell = row.createCell(0);
            cell.setCellValue(saleDate);
            cell = row.createCell(1);
            cell.setCellValue(spareName);
            cell = row.createCell(2);
            cell.setCellValue(quantity);
            cell = row.createCell(3);
            cell.setCellValue(price);
            inputStream.close();
            FileOutputStream outputStream = new FileOutputStream(System.getProperty("user.dir") + "\\src\\main\\resources\\Data\\daysSale.xlsx");
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();
        } catch (IOException | EncryptedDocumentException ex) {
            assert ex != null;
            ex.printStackTrace();
            System.out.println("Error Occured while adding data");
        }
    }

    public void updateQuantity(String spareName, int quantity) throws IOException {
        try {
            FileInputStream fis = new FileInputStream(System.getProperty("user.dir") + "\\src\\main\\resources\\Data\\SparesList.xlsx");   //obtaining bytes from the file
            Workbook workbook = WorkbookFactory.create(fis);
            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> itr = sheet.iterator();    //iterating over excel file
            int i = 0;
            while (itr.hasNext()) {
                Row row = itr.next();
                if (row.getCell(0).toString().equalsIgnoreCase(spareName)) {
                    Cell cell2Update = sheet.getRow(i).getCell(2);
                    cell2Update.setCellValue(quantity);
                }
                i++;
            }

            fis.close();
            FileOutputStream outputStream = new FileOutputStream(System.getProperty("user.dir") + "\\src\\main\\resources\\Data\\SparesList.xlsx");
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();
        } catch (IOException | EncryptedDocumentException ex) {
            assert ex != null;
            ex.printStackTrace();
            System.out.println("Error Occured while adding data");
        }
    }

    public void closeCounter() throws IOException {
        getTodaysSale();
        String[] availableSpares = getAvailableSparesName();
        int[] averageSalesValue = calculateAverageSale(availableSpares);
        updateAverageSalesInSheet(averageSalesValue, availableSpares);
        orderSpares();
    }

    public int getTodaysSale() throws IOException {
        int avg = 0;
        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
        FileInputStream inputStream = new FileInputStream(new File(System.getProperty("user.dir") + "\\src\\main\\resources\\Data\\daysSale.xlsx"));
        Workbook workbook = WorkbookFactory.create(inputStream);
        Sheet sheet = workbook.getSheetAt(0);
        Iterator<Row> itr = sheet.iterator();    //iterating over excel file
        int i = 0, price = 0;
        while (itr.hasNext()) {
            Row row = itr.next();
            Date date = new Date();
            String saleDate = formatter.format(date);
            if (row.getCell(0).toString().equals(saleDate)) {
                price = price + Integer.parseInt(row.getCell(3).toString().split("\\.")[0]);
            }
            i++;
        }
        System.out.println("Todays total Sale: " + price);
        avg = price / i;
        return avg;
    }

    public String[] getAvailableSparesName() throws IOException {
        String[] spareDetails = new String[10];
        String[] identifier = new String[4];
        try {
            String path = System.getProperty("user.dir") + "\\src\\main\\resources\\Data\\SparesList.xlsx";
            System.out.println("PATH: " + path);
            File file = new File(path);   //creating a new file instance
            FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object
            int rowCount = sheet.getLastRowNum();
            spareDetails = new String[rowCount];
            Iterator<Row> itr = sheet.iterator();    //iterating over excel file
            int i = 0;
            while (itr.hasNext()) {
                Row row = itr.next();
                Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column
                if (i != 0)
                    spareDetails[i - 1] = String.valueOf(row.getCell(0));
                i++;
            }
            fis.close();
            wb.close();
        } catch (Exception e) {
            System.out.println("Error reading file please try again");
        }
        return spareDetails;
    }

    public int[] calculateAverageSale(String[] sparesList) {
        int sale = 0;
        int[] saleQuantity = new int[sparesList.length];
        System.out.println(Arrays.toString(sparesList));
        try {
            SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
            Date date = new Date();
            String saleDate = formatter.format(date);
            int i = 0;
            for (String s : sparesList) {
                sale = 0;
                FileInputStream inputStream = new FileInputStream(new File(System.getProperty("user.dir") + "\\src\\main\\resources\\Data\\daysSale.xlsx"));
                Workbook workbook = WorkbookFactory.create(inputStream);
                Sheet sheet = workbook.getSheetAt(0);
                Iterator<Row> itr = sheet.iterator();
                ;
                while (itr.hasNext()) {
                    Row row = itr.next();
                    Iterator<Cell> cellIterator = row.cellIterator();
                    if (row.getCell(1).toString().equalsIgnoreCase(s) && row.getCell(0).toString().equals(saleDate)) {
                        sale = (int) (sale + row.getCell(2).getNumericCellValue());
                    }
                }
                saleQuantity[i] = sale;
                inputStream.close();
                workbook.close();
                i++;
            }
        } catch (Exception e) {
            System.out.println("Error occured reading data please try again");
        }
        return saleQuantity;
    }

    public void updateAverageSalesInSheet(int[] averageDetails, String[] spareDetails) throws IOException {
        for (int i = 0; i < spareDetails.length; i++) {
            String path = System.getProperty("user.dir") + "\\src\\main\\resources\\Data\\SparesList.xlsx";
            File file = new File(path);   //creating a new file instance
            FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file
            Workbook wb = WorkbookFactory.create(fis);
            Sheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object
            Iterator<Row> itr = sheet.iterator();    //iterating over excel file
            int j = 0;
            while (itr.hasNext()) {
                Row row = itr.next();
                if (row.getCell(0).toString().equalsIgnoreCase(spareDetails[i])) {
                    Cell cell = sheet.getRow(j).createCell(5);
                    cell.setCellValue(averageDetails[i]);
                    System.out.println(averageDetails[i]);
                }
                j++;
            }
            fis.close();
            FileOutputStream outputStream = new FileOutputStream(System.getProperty("user.dir") + "\\src\\main\\resources\\Data\\SparesList.xlsx");
            wb.write(outputStream);
            wb.close();
            outputStream.close();
        }
    }

    public void orderSpares() throws IOException {
        String path = System.getProperty("user.dir") + "\\src\\main\\resources\\Data\\SparesList.xlsx";
        File file = new File(path);   //creating a new file instance
        FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file
        Workbook wb = WorkbookFactory.create(fis);
        Sheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object
        Iterator<Row> itr = sheet.iterator();    //iterating over excel file
        int j = 0;
        while (itr.hasNext()) {
            Row row = itr.next();
            if (j > 0) {
                int quantity = Integer.parseInt(row.getCell(2).toString().split("\\.")[0]);
                int average = Integer.parseInt(row.getCell(5).toString().split("\\.")[0]);
                if (quantity <= average) {
                    Cell cell = sheet.getRow(j).getCell(2);
                    cell.setCellValue(quantity + average);
                    System.out.println("placed " + (quantity + average) + ": new order for spare " + sheet.getRow(j).getCell(0));
                }
            }
            j++;

        }
        fis.close();
        FileOutputStream outputStream = new FileOutputStream(System.getProperty("user.dir") + "\\src\\main\\resources\\Data\\SparesList.xlsx");
        wb.write(outputStream);
        wb.close();
        outputStream.close();
    }

    public int[] calculateAverageWeeklySales(String[] sparesList) throws IOException, ParseException {
        int sale = 0;
        int[] saleQuantity = new int[sparesList.length];
        System.out.println(Arrays.toString(sparesList));
//        try {
        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
        Date date = new Date();
        String saleDate = formatter.format(date);
        Calendar cal = Calendar.getInstance();
        cal.setTime(formatter.parse(String.valueOf(saleDate)));
        cal.add(Calendar.DATE, -7);
        String minus7days = formatter.format(cal.getTime());
        cal.add(Calendar.DATE, -6);
        String minus6days = formatter.format(cal.getTime());
        cal.add(Calendar.DATE, -5);
        String minus5days = formatter.format(cal.getTime());
        cal.add(Calendar.DATE, -4);
        String minus4days = formatter.format(cal.getTime());
        cal.add(Calendar.DATE, -3);
        String minus3days = formatter.format(cal.getTime());
        cal.add(Calendar.DATE, -2);
        String minus2days = formatter.format(cal.getTime());
        cal.add(Calendar.DATE, -1);
        String minus1days = formatter.format(cal.getTime());
        int i = 0;
        for (String s : sparesList) {
            sale = 0;
            FileInputStream inputStream = new FileInputStream(new File(System.getProperty("user.dir") + "\\src\\main\\resources\\Data\\daysSale.xlsx"));
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> itr = sheet.iterator();
            while (itr.hasNext()) {
                Row row = itr.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                if (row.getCell(1).toString().equalsIgnoreCase(s) &&
                        (row.getCell(0).toString().equals(saleDate)) ||
                        (row.getCell(0).toString().equals(minus1days)) ||
                        (row.getCell(0).toString().equals(minus2days)) ||
                        (row.getCell(0).toString().equals(minus3days)) ||
                        (row.getCell(0).toString().equals(minus4days)) ||
                        (row.getCell(0).toString().equals(minus5days)) ||
                        (row.getCell(0).toString().equals(minus6days)) ||
                        (row.getCell(0).toString().equals(minus7days))) {
                    sale = (int) (sale + row.getCell(2).getNumericCellValue());
                }
            }
            saleQuantity[i] = sale;
            inputStream.close();
            workbook.close();
            i++;
        }
//        } catch (Exception e) {
//            System.out.println("Error occured reading data please try again");
//        }
        return saleQuantity;
    }
}