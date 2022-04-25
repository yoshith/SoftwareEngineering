import java.io.IOException;
import java.text.ParseException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Scanner;

public class Main {
    static HashMap<String, String> data = new HashMap<String, String>();
    public static void menu() throws IOException, InterruptedException, ParseException {
        ExcelHelper excel = new ExcelHelper();
        Scanner sc = new Scanner(System.in);
        System.out.println("""
                Please enter the numeric from the below Menu:\s
                1: get Spares Data\s
                2: add Spares Data\s
                3: buy Spare\s
                4: close counter\s
                5: weekly Sales Report\s
                6: close program
                """);
        switch (sc.next()) {
            case "1":
                System.out.println("enter Spare Name without Space: ");
                String spareName = sc.next();
                System.out.println(spareName);
                data = excel.getSpareData(spareName);
                System.out.println(data.keySet());
                System.out.println(data.values());
                menu();
                break;
            case "2":
                String[] data = new String[5];
                String spareAdded = "";
                for(int i=0; i<= data.length;i++){
                    if(i==0){
                        System.out.println("Enter Spare Name to Add: (please add name without space)");
                        spareAdded = sc.next();
                        data[i] = spareAdded;
                    }else if(i==1){
                        System.out.println("Enter Vendor Name to Add: (please add data without space)");
                        data[i]=sc.next();
                    }else if(i==2){
                        System.out.println("Enter Quantity of spare to Add: ");
                        data[i]=sc.next();
                    }else if(i==3){
                        System.out.println("Enter price of Spare to Add: ");
                        data[i]=sc.next();
                    }else if(i==4){
                        System.out.println("Enter address: (please add Address without space)");
                        data[i]=sc.next();
                    }
                }
                excel.addSparesData(data);
                System.out.println("spare Data added:" + spareAdded);
                Thread.sleep(5000);
                excel.getSpareData(spareAdded.trim());
                menu();
                break;
            case "3":
                System.out.println("Please select spare name from below List: ");
                System.out.println(Arrays.toString(excel.getAvailableSparesName()));
                excel.getSpareDataForSales(sc.next());
                menu();
                break;
            case "4":
                System.out.println("Closing counter......");
                excel.closeCounter();
                menu();
                break;
            case "5":
                System.out.println("weekly Sales Report: ");
                String[] availableSpares = excel.getAvailableSparesName();
                System.out.println(Arrays.toString(availableSpares));
                System.out.println(Arrays.toString(excel.calculateAverageWeeklySales(availableSpares)));
                menu();
                break;
            case "6":
                break;
            default:
                System.out.println("Please select correct value: ");
                menu();
                break;
        }
    }

    public static void main(String[] args) throws IOException, InterruptedException, ParseException {
        menu();
    }


}