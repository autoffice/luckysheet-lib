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

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.databind.DeserializationFeature;
import com.fasterxml.jackson.databind.ObjectMapper;
import lombok.AccessLevel;
import lombok.NoArgsConstructor;

import java.io.File;

/**
 * JSON 序列化/反序列化工具类.
 *
 * <p>基于 Jackson ObjectMapper, 配置为忽略 null 值和未知属性.</p>
 */
@NoArgsConstructor(access = AccessLevel.PRIVATE)
public final class JSONUtil {
    private static final ObjectMapper mapper = new ObjectMapper();

    static {
        mapper.setSerializationInclusion(JsonInclude.Include.NON_NULL);
        mapper.configure(DeserializationFeature.FAIL_ON_UNKNOWN_PROPERTIES, false);
    }

    /**
     * 将 JSON 文件反序列化为指定类型的对象.
     *
     * @param file   JSON 文件
     * @param entity 目标类型
     * @param <T>    目标类型泛型
     * @return 反序列化后的对象
     */
    public static <T> T asEntity(File file, Class<T> entity) {
        try {
            return mapper.readValue(file, entity);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 将对象序列化为 JSON 字符串.
     *
     * @param entity 待序列化的对象
     * @return JSON 字符串
     */
    public static String asJSONString(Object entity) {
        try {
            return mapper.writeValueAsString(entity);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}
