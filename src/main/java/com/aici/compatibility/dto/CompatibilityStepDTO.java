package com.aici.compatibility.dto;

import lombok.Data;
import lombok.experimental.Accessors;

/**
 * 兼容性测试 - 步骤 - 数据体
 *
 * @author <a href="mailto:472546172@qq.com">张天赐</a>
 */
@Data
@Accessors(chain = true)
public class CompatibilityStepDTO {
    /**
     * 步骤名称
     */
    private String stepName;
    /**
     * 前提条件
     */
    private String precondition;
    /**
     * 预期结果
     */
    private String expectedResult;
    /**
     * 步骤描述
     */
    private String stepDescription;
}
