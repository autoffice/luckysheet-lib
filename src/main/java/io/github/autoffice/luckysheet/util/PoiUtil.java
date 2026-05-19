/*
 * Copyright © 2025 AutOffice (hello.aldis@qq.com)
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package io.github.autoffice.luckysheet.util;

import lombok.AccessLevel;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Apache POI 操作的辅助工具类.
 *
 * <p>提供基于 0 起始的行号取值、最大列数计算等帮助方法.</p>
 */
@NoArgsConstructor(access = AccessLevel.PRIVATE)
public final class PoiUtil {
    /**
     * 获取工作表中第一个非空行的行号 (从 0 开始, 即使 POI 返回 -1 也会修正为 0).
     *
     * @param sheet 工作表
     * @return 起始行号 (大于等于 0)
     */
    public static int getFirstRowBase0(Sheet sheet) {
        int firstRowNum = sheet.getFirstRowNum();
        return Math.max(firstRowNum, 0);
    }

    /**
     * 计算工作表中所有行的最大列数.
     *
     * @param sheet 工作表
     * @return 最大列数 (即列号上界, 用于遍历)
     */
    public static short getMaxColNum(Sheet sheet) {
        short maxColNum = 0;
        for (int i = getFirstRowBase0(sheet); i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                maxColNum = (short) Math.max(maxColNum, row.getLastCellNum());
            }
        }

        return maxColNum;
    }
}
