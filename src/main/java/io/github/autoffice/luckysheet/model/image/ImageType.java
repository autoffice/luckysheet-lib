package io.github.autoffice.luckysheet.model.image;

import com.fasterxml.jackson.annotation.JsonValue;
import io.github.autoffice.luckysheet.util.Util;
import lombok.AllArgsConstructor;
import lombok.Getter;
import org.apache.poi.ss.usermodel.ClientAnchor;

import java.util.Arrays;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;

@AllArgsConstructor
public enum ImageType {
    /**
     * 移动并调整单元格大小
     */
    MOVE_AND_RESIZE(1, ClientAnchor.AnchorType.MOVE_AND_RESIZE),
    /**
     * 移动并且不调整单元格的大小
     */
    MOVE_DONT_RESIZE(2, ClientAnchor.AnchorType.MOVE_DONT_RESIZE),
    /**
     * 不要移动单元格并调整大小
     */
    DONT_MOVE_AND_RESIZE(3, ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);

    @JsonValue
    private final Integer lsValue;

    @Getter
    private final ClientAnchor.AnchorType poiValue;

    private static final Map<ClientAnchor.AnchorType, ImageType> TYPES = Arrays.stream(values())
            .collect(Collectors.toMap(ImageType::getPoiValue, Function.identity()));

    public static ImageType of(ClientAnchor.AnchorType anchorType) {
        ImageType imageType = TYPES.get(anchorType);
        return Util.requireNonNullElse(imageType, MOVE_AND_RESIZE);
    }
}
