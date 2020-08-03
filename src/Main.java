import pakahi.excel.*;
import java.nio.file.Path;
import java.nio.file.Paths;



//----------------------------------------------------------------------------------------------------
public class Main {

    //------------------------------------------------------------------------------------------------
    public static void main(String[] args) {

        String path = System.getProperty("user.dir");
        String templateName = "template";
        Path template = Paths.get(path, "templates", templateName);

        IReport report = new InvoiceData();

        try(Workbook workbook = new Workbook(template, report)) {

            // ファイルとして保存（パスワードなし）
            workbook.save(Paths.get(path, templateName + ".xlsx"));
/*
            // ファイルとして保存（パスワードあり）
            workbook.save(Paths.get(path, templateName + "-password.xlsx"), "123");

            // バイト配列として生成（パスワードなし）
            byte[] bytes1 = workbook.save();
            try(FileOutputStream f = new FileOutputStream(Paths.get(folder, name + "-binary.xlsx").toFile())) {
                f.write(bytes1);
            } catch(Exception ex) {
            }

            // バイト配列として生成（パスワードあり）
            byte[] bytes2 = workbook.save("123");
            System.out.println("bytes2=" + bytes2.length);
            try(FileOutputStream f = new FileOutputStream(Paths.get(folder, name + "-binary-password.xlsx").toFile())) {
                f.write(bytes2);
            } catch(Exception ex) {
            }
*/
        }

    }
    //------------------------------------------------------------------------------------------------

}
