package com.aspose.spreadsheeteditor;

import java.awt.Color;
import javax.faces.view.ViewScoped;
import javax.inject.Named;

/**
 *
 * @author Saqib Masood
 */
@Named(value = "formatting")
@ViewScoped
public class FormattingService {

    private boolean boldOptionEnabled;
    private boolean italicOptionEnabled;
    private boolean underlineOptionEnabled;
    private String fontSelectionOption;
    private int fontSizeOption;
    private String alignSelectionOption;
    private Color fontColorSelectionOption;

    public boolean isBoldOptionEnabled() {
        return boldOptionEnabled;
    }

    public void setBoldOptionEnabled(boolean boldOptionEnabled) {
        this.boldOptionEnabled = boldOptionEnabled;
    }

    public boolean isItalicOptionEnabled() {
        return italicOptionEnabled;
    }

    public void setItalicOptionEnabled(boolean italicOptionEnabled) {
        this.italicOptionEnabled = italicOptionEnabled;
    }

    public boolean isUnderlineOptionEnabled() {
        return underlineOptionEnabled;
    }

    public void setUnderlineOptionEnabled(boolean underlineOptionEnabled) {
        this.underlineOptionEnabled = underlineOptionEnabled;
    }

    public String getFontSelectionOption() {
        return fontSelectionOption;
    }

    public void setFontSelectionOption(String fontSelectionOption) {
        this.fontSelectionOption = fontSelectionOption;
    }

    public int getFontSizeOption() {
        return fontSizeOption;
    }

    public void setFontSizeOption(int fontSizeOption) {
        this.fontSizeOption = fontSizeOption;
    }

    public String getAlignSelectionOption() {
        return alignSelectionOption;
    }

    public void setAlignSelectionOption(String alignSelectionOption) {
        this.alignSelectionOption = alignSelectionOption;
    }

    public Color getFontColorSelectionOption() {
        return this.fontColorSelectionOption;
    }

    public void setFontColorSelectionOption(Color c) {
        this.fontColorSelectionOption = c;
    }

}
