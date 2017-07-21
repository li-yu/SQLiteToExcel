package com.liyu.sqlitetoexcel;

import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

/**
 * Created by liyu on 2017/7/20.
 */

public class SecurityUtil {

    /**
     * Encrypt a file, support .xls file only
     *
     * @param file
     * @param encryptKey
     * @throws Exception
     */
    public static void EncryptFile(File file, String encryptKey) throws Exception {
        FileInputStream fileInput = new FileInputStream(file.getPath());
        BufferedInputStream bufferInput = new BufferedInputStream(fileInput);
        POIFSFileSystem poiFileSystem = new POIFSFileSystem(bufferInput);
        Biff8EncryptionKey.setCurrentUserPassword(encryptKey);
        HSSFWorkbook workbook = new HSSFWorkbook(poiFileSystem, true);
        FileOutputStream fileOut = new FileOutputStream(file.getPath());
        workbook.writeProtectWorkbook(Biff8EncryptionKey.getCurrentUserPassword(), "");
        workbook.write(fileOut);
        bufferInput.close();
        fileOut.close();
    }

}
