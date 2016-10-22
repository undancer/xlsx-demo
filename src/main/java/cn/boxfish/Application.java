package cn.boxfish;

import com.google.common.base.Stopwatch;
import com.google.common.collect.Table;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.xml.sax.SAXException;

import java.io.IOException;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.concurrent.TimeUnit;

/**
 * Created by undancer on 2016/10/20.
 */
public class Application {

    public static void main(String[] args) throws IOException {
//        String path = "/Users/undancer/Documents/boxfish/1.文档备份/数据表/总表/Aperture图片例句整理 20130411TMY.xlsx";

        Path start = Paths.get("/Users/undancer/Documents/boxfish/");
        Stopwatch timer = Stopwatch.createStarted();

        Files.walkFileTree(start, new SimpleFileVisitor<Path>() {

            public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) throws IOException {

                String filename = file.getFileName().toString();

                if (!StringUtils.startsWith(filename, "~$") && StringUtils.endsWith(filename, ".xlsx")) {
                    try {
                        System.out.println(file);
                        printFile(file.toString());
                    } catch (SAXException e) {
                        e.printStackTrace();
                    } catch (OpenXML4JException e) {
                        e.printStackTrace();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }

                return FileVisitResult.CONTINUE;
            }
        });
        System.err.println(timer.elapsed(TimeUnit.MILLISECONDS));
    }

    public static void printFile(String file) throws IOException, SAXException, OpenXML4JException {
        Reader reader = new Reader(file);

        for (Reader.Sheet sheet : reader.getSheets()) {
            Table<Integer, Integer, String> table = reader.getSheet(sheet.getRelId());

            for (Table.Cell<Integer, Integer, String> cell : table.cellSet()) {
                System.out.printf("%s - %s - [%d,%d] %s \n", file, sheet.getName(), cell.getRowKey(), cell.getColumnKey(), cell.getValue());

            }
        }

    }


}
