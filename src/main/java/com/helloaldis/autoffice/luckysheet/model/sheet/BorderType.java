package com.helloaldis.autoffice.luckysheet.model.sheet;

import com.fasterxml.jackson.annotation.JsonValue;
import lombok.AllArgsConstructor;

@AllArgsConstructor
public enum BorderType {
    /**
     *
     */
    LEFT("border-left"),
    /**
     *
     */
    RIGHT("border-right"),
    /**
     *
     */
    TOP("border-top"),
    /**
     *
     */
    BOTTOM("border-bottom"),
    /**
     *
     */
    ALL("border-all"),
    /**
     *
     */
    OUTSIDE("border-outside"),
    /**
     *
     */
    INSIDE("border-inside"),
    /**
     *
     */
    HORIZONTAL("border-horizontal"),
    /**
     *
     */
    VERTICAL("border-vertical"),
    /**
     *
     */
    NONE("border-none");

    @JsonValue
    private final String lsValue;
}
