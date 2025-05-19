package io.github.autoffice.luckysheet.model;

import io.github.autoffice.luckysheet.model.sheet.LuckySheet;
import lombok.Data;

import java.util.List;

@Data
public class LuckyFile {
    private LuckyFileInfo info;
    private List<LuckySheet> sheets;
}
