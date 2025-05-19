package io.github.autoffice.luckysheet.model.image;

import com.fasterxml.jackson.annotation.JsonValue;
import lombok.AllArgsConstructor;

@AllArgsConstructor
public enum ImageBorderType {
    /**
     * 实线
     */
    SOLID("solid"),
    /**
     * 虚线
     */
    DASHED("dashed"),
    /**
     * 点状
     */
    DOTTED("dotted"),
    /**
     * 双线
     */
    DOUBLE("double");

    @JsonValue
    private final String lsValue;
}
