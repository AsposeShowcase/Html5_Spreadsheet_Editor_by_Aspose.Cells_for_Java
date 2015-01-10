package com.aspose.spreadsheeteditor;

import com.aspose.cells.CellsException;
import com.aspose.cells.License;
import java.io.IOException;
import java.io.InputStream;

/**
 *
 * @author Saqib Masood
 */
public class AsposeLicense {

    public static final String FILE_NAME = "Aspose.Total.Java.lic";

    public static void load() {
    }

    static {
        try (InputStream i = AsposeLicense.class.getResourceAsStream(FILE_NAME)) {
            new License().setLicense(i);
        } catch (IOException x) {
            System.out.println("Error occured while loading license: " + x.getMessage());
        } catch (CellsException x) {
            System.out.println("License error: " + x.getMessage());
        }
    }
}
