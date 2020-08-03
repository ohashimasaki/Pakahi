package pakahi.excel;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import javax.xml.namespace.NamespaceContext;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.*;
import java.lang.reflect.Field;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;



//----------------------------------------------------------------------------------------------------
/**
 * 展開済みの SpreadsheetML 構造を Excel ファイルとして操作するクラスです。
 */
public class Workbook implements AutoCloseable {

    SharedStrings sharedStrings;
    Path tmp;
    HashMap<String, String> sheets = new HashMap<>();
    IReport report;
    NamespaceContext context;


    //------------------------------------------------------------------------------------------------
    /**
    * 展開済みの SpreadsheetML 構造を Excel ファイルとして操作するクラスのコンストラクタです。
     * @param template テンプレートとして使用する展開済み SpreadsheetML フォルダのパス
     * @param report テンプレートに挿入する値としてのデータオブジェクト
     */
    public Workbook(Path template, IReport report) {

        context = new ExcelNamespaceContext();

        // データ
        this.report = report;

        try {
            // テンプレートを一時フォルダとしてコピー
            tmp = Files.createTempDirectory(null);
            copy(template, tmp);

            // ワークシートのリストを読み込み
            sheets.putAll(getSheets(tmp));

            // 共有文字列の読み込み
            sharedStrings = new SharedStrings(tmp);

            // データの書き込み
            render(tmp);

        } catch(Exception ex) {
            System.out.println(Arrays.toString(ex.getStackTrace()));
        }

    }
    //------------------------------------------------------------------------------------------------
    /**
     * クローズ時に自動的に一時フォルダを削除します。
     */
    @SuppressWarnings("ResultOfMethodCallIgnored")
    public void close() {

        if(tmp == null || ! Files.exists(tmp)) {
            return;
        }

        try {
            // Excel ファイル生成用の一時フォルダを削除
            Files.walk(tmp).sorted(Comparator.reverseOrder()).map(Path::toFile).forEach(File::delete);
        } catch(Exception ex) {
            System.out.println(Arrays.toString(ex.getStackTrace()));
        }

    }
    //------------------------------------------------------------------------------------------------
    private void render(Path tmp) {

        if(tmp == null || ! Files.exists(tmp)) {
            return;
        }

        // データをパースしてシートごとに格納
        HashMap<String, HashMap<String, Object>> itemsPerSheet = new HashMap<>() {{
            putAll(parseFields());
        }};

        // シートごとに書き込む
        try {
            for(String sheetId : itemsPerSheet.keySet()) {
                Worksheet worksheet = new Worksheet(tmp, sheetId, sharedStrings);
                HashMap<String, Object> items = itemsPerSheet.get(sheetId);

                for(String address : items.keySet()) {
                    Object value = items.get(address);
                    if( ! (value instanceof List<?>) && ! (value instanceof Map<?, ?>)) {
                        setCellValue(worksheet, address, value);
                    }
                }
                for(String address : items.keySet()) {
                    Object value = items.get(address);
                    if(value instanceof List<?>) {
                        insertRows(worksheet, address, value);
                    }
                }

                worksheet.save();
            }

        } catch(Exception ex) {
            System.out.println(Arrays.toString(ex.getStackTrace()));
        }

    }
    //------------------------------------------------------------------------------------------------
    private HashMap<String, HashMap<String, Object>> parseFields() {

        HashMap<String, HashMap<String, Object>> itemsPerSheet = new HashMap<>();

        try {
            for(Field f : report.getClass().getDeclaredFields()) {
                if(f.isAnnotationPresent(Range.class)) {
                    // シート名を表示上のものから内部 ID に変換
                    String sheetName = f.getAnnotation(Range.class).sheet();
                    if(sheets.containsKey(sheetName)) {
                        String sheetId = sheets.get(sheetName);
                        if( ! itemsPerSheet.containsKey(sheetId)) {
                            itemsPerSheet.put(sheetId, new HashMap<>());
                        }
                        String address = f.getAnnotation(Range.class).range();
                        Object value = f.get(report);
                        itemsPerSheet.get(sheetId).put(address, value);
                    }
                }
            }
        } catch(Exception ex) {
            System.out.println(Arrays.toString(ex.getStackTrace()));
        }

        return itemsPerSheet;

    }
    //------------------------------------------------------------------------------------------------
    @SuppressWarnings("unchecked")
    private void insertRows(Worksheet worksheet, String address, Object value) {

        try {
            List<IReport> items = (List<IReport>)value;

            if(items.size() == 0) {
                return;
            }

            Address a = new Address(address);
            int r = a.getRow();
            int c = a.getColumn();

            worksheet.insertRows(a.getRow(), items.size());

            for(int i = 0; i < items.size(); i++) {
                IReport e = items.get(i);
                for(Field f : e.getClass().getDeclaredFields()) {
                    if(f.isAnnotationPresent(RangeR1C1.class)) {
                        int columnOffset = f.getAnnotation(RangeR1C1.class).column() - 1;
                        setCellValue(worksheet, a.set(r + i, c + columnOffset).getA1(), f.get(e));
                    }
                }
            }
        } catch(Exception ex) {
            System.out.println(Arrays.toString(ex.getStackTrace()));
        }

    }
    //------------------------------------------------------------------------------------------------
    private void setCellValue(Worksheet worksheet, String address, Object value) {

        Class<?> type = value.getClass();

        if(type.equals(Integer.class)) {
            worksheet.setCellValue(address, (int)value);

        } else if(type.equals(Double.class)) {
            worksheet.setCellValue(address, (double)value);

        } else if(type.equals(String.class)) {
            worksheet.setCellValue(address, (String)value);

        } else if(type.equals(Long.class)) {
            worksheet.setCellValue(address, (long)value);

        } else if(type.equals(Float.class)) {
            worksheet.setCellValue(address, (float)value);

        } else if(type.equals(Byte.class)) {
            worksheet.setCellValue(address, (byte)value);

        } else if(type.equals(Short.class)) {
            worksheet.setCellValue(address, (short)value);

        } else if(type.equals(Boolean.class)) {
            worksheet.setCellValue(address, (boolean)value);

        } else if(type.equals(LocalDateTime.class)) {
            worksheet.setCellValue(address, (LocalDateTime)value);

        } else if(type.equals(LocalDate.class)) {
            worksheet.setCellValue(address, (LocalDate)value);

        } else if(type.equals(LocalTime.class)) {
            worksheet.setCellValue(address, (LocalTime)value);
        }

    }
    //------------------------------------------------------------------------------------------------
    /**
     * Excel ファイルをバイト配列として返します。パスワードによる暗号化はされません。
     * @return Excelファイルのバイト配列
     */
    public byte[] save() {

        return save((String)null);

    }
    //------------------------------------------------------------------------------------------------
    /**
     * パスワードにより暗号化された Excel ファイルをバイト配列として返します。
     * @param password パスワード
     * @return 暗号化された Excel ファイルのバイト配列
     */
    public byte[] save(String password) {

        sharedStrings.save();
        byte[] bytes = zip(tmp);

        if(password != null && ! password.isEmpty()) {
            return encrypt(bytes, password);
        } else {
            try(OPCPackage opc = OPCPackage.open(new ByteArrayInputStream(bytes));
                ByteArrayOutputStream output = new ByteArrayOutputStream()) {
                opc.save(output);
                return output.toByteArray();
            } catch(Exception ex) {
                System.out.println(Arrays.toString(ex.getStackTrace()));
                return null;
            }
        }

    }
    //------------------------------------------------------------------------------------------------
    /**
     * Excel ファイルを保存します。パスワードによる暗号化はされません。
     * @param path 保存するファイルのパス
     */
    public void save(Path path) {

        save(path, (String)null);

    }
    //------------------------------------------------------------------------------------------------
    /**
     * パスワードにより暗号化された Excel ファイルを保存します。
     * @param path 保存するファイルのパス
     * @param password パスワード
     */
    public void save(Path path, String password) {

        sharedStrings.save();

        zip(tmp, path);

        if(password != null && ! password.isEmpty()) {
            encrypt(path, password);
        }

    }
    //------------------------------------------------------------------------------------------------
    private byte[] encrypt(byte[] bytes, String password) {

        byte[] encryptedBytes = null;

        try(POIFSFileSystem poifs = new POIFSFileSystem()) {
            Encryptor encryptor = new EncryptionInfo(EncryptionMode.agile).getEncryptor();
            encryptor.confirmPassword(password);

            try(OPCPackage opc = OPCPackage.open(new ByteArrayInputStream(bytes));
                OutputStream output = encryptor.getDataStream(poifs)) {
                opc.save(output);
            } catch(Exception ex) {
                System.out.println(Arrays.toString(ex.getStackTrace()));
                return null;
            }

            try(ByteArrayOutputStream output = new ByteArrayOutputStream()) {
                poifs.writeFilesystem(output);
                encryptedBytes = output.toByteArray().clone();
            } catch(Exception ex) {
                System.out.println(Arrays.toString(ex.getStackTrace()));
                return null;
            }

        } catch(Exception ex) {
            System.out.println(Arrays.toString(ex.getStackTrace()));
        }

       return encryptedBytes;

    }
    //------------------------------------------------------------------------------------------------
    private void encrypt(Path path, String password) {

        File file = path.toFile();

        try(POIFSFileSystem poifs = new POIFSFileSystem()) {
            Encryptor encryptor = new EncryptionInfo(EncryptionMode.agile).getEncryptor();
            encryptor.confirmPassword(password);

            try(OPCPackage opc = OPCPackage.open(file, PackageAccess.READ_WRITE);
                OutputStream output = encryptor.getDataStream(poifs)) {
                opc.save(output);
            } catch(Exception ex) {
                System.out.println(Arrays.toString(ex.getStackTrace()));
            }

            try(FileOutputStream output = new FileOutputStream(file)) {
                poifs.writeFilesystem(output);
            } catch(Exception ex) {
                System.out.println(Arrays.toString(ex.getStackTrace()));
            }

        } catch(Exception ex) {
            System.out.println(Arrays.toString(ex.getStackTrace()));
        }

    }
    //------------------------------------------------------------------------------------------------
    private byte[] zip(Path source) {

        byte[] bytes = null;

        try(ByteArrayOutputStream output = new ByteArrayOutputStream();
            ZipOutputStream zip = new ZipOutputStream(output)) {
            Files.walk(source).filter(path -> ! Files.isDirectory(path)).forEach(path -> {
                try {
                    ZipEntry e = new ZipEntry(source.relativize(path).toString().replace("\\", "/"));
                    zip.putNextEntry(e);
                    Files.copy(path, zip);
                    zip.closeEntry();
                } catch(Exception ex) {
                    System.out.println(Arrays.toString(ex.getStackTrace()));
                }
            });

            bytes = output.toByteArray().clone();

        } catch(Exception ex) {
            System.out.println(Arrays.toString(ex.getStackTrace()));
        }

        return bytes;

    }
    //------------------------------------------------------------------------------------------------
    private void zip(Path source, Path target) {

        try(FileOutputStream output = new FileOutputStream(target.toFile());
            ZipOutputStream zip = new ZipOutputStream(output)) {
            Files.walk(source).filter(path -> ! Files.isDirectory(path)).forEach(path -> {
                try {
                    zip.putNextEntry(new ZipEntry(source.relativize(path).toString().replace("\\", "/")));
                    Files.copy(path, zip);
                    zip.closeEntry();
                } catch(Exception ex) {
                    System.out.println(Arrays.toString(ex.getStackTrace()));
                }
            });

        } catch(Exception ex) {
            System.out.println(Arrays.toString(ex.getStackTrace()));
        }

    }
    //------------------------------------------------------------------------------------------------
    private void copy(Path from, Path to) {

        try {
            Files.walk(from).forEach(path -> {
                Path f = to.resolve(from.relativize(path));
                if( ! Files.exists(f)) {
                    if(Files.isDirectory(path)) {
                        try {
                            Files.createDirectory(f);
                        } catch(Exception ex) {
                            System.out.println(Arrays.toString(ex.getStackTrace()));
                        }
                    } else {
                        try {
                            Files.copy(path, f);
                        } catch(Exception ex) {
                            System.out.println(Arrays.toString(ex.getStackTrace()));
                        }
                    }
                }
            });
        } catch(Exception ex) {
            System.out.println(Arrays.toString(ex.getStackTrace()));
        }

    }
    //------------------------------------------------------------------------------------------------
    private HashMap<String, String> getSheets(Path tmp) {

        HashMap<String, String> sheets = new HashMap<>();

        File file = Paths.get(tmp.toString(), "xl/workbook.xml").toFile();

        if( ! file.exists()) {
            return sheets;
        }

        try {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            Document xml = factory.newDocumentBuilder().parse(file);

            List<Element> items = XPath.selectNodes(xml, "/x:workbook/x:sheets/x:sheet");

            if(items != null) {
                String uri = context.getNamespaceURI("o");
                for(Element e : items) {
                    String id = e.getAttributeNS(uri, "id");
                    String name = e.getAttribute("name");
                    sheets.put(name, id.replace("rId", "sheet"));
                }
            }

            return sheets;

        } catch(Exception ex) {
            System.out.println(Arrays.toString(ex.getStackTrace()));
            return sheets;
        }

    }
    //------------------------------------------------------------------------------------------------

}
//----------------------------------------------------------------------------------------------------
