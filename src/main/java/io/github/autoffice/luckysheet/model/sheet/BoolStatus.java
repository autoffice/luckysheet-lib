package io.github.autoffice.luckysheet.model.sheet;

import com.fasterxml.jackson.annotation.JsonValue;
import lombok.AllArgsConstructor;
import lombok.Getter;

@AllArgsConstructor
public enum BoolStatus {
    /**
     * false 0
     */
    FALSE(0, false),
    /**
     * true 1
     */
    TRUE(1, true);

    @JsonValue
    private final Integer lsValue;

    @Getter
    private final boolean poiValue;

    public static BoolStatus of(boolean poiValue) {
        return poiValue ? TRUE : FALSE;
    }
}
