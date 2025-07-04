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

@NoArgsConstructor(access = AccessLevel.PRIVATE)
public final class PoiUtil {
    public static int getFirstRowBase0(Sheet sheet) {
        int firstRowNum = sheet.getFirstRowNum();
        return Math.max(firstRowNum, 0);
    }

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
