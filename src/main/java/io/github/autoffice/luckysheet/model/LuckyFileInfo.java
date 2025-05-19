package io.github.autoffice.luckysheet.model;

import lombok.Data;

/**
 * Excel文件基本信息
 */
@Data
public class LuckyFileInfo {
    /**
     * 文件名
     */
    private String name;
    /**
     * 创建人
     */
    private String creator;
    /**
     * 最新修改人
     */
    private String lastmodifiedby;
    /**
     * 创建时间
     */
    private String createdTime;
    /**
     * 最新修改时间
     */
    private String modifiedTime;
    /**
     * 公司信息
     */
    private String company;
    /**
     * excel version
     */
    private String appversion;
}
