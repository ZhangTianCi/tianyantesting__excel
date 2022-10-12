package com.aici.compatibility.dto;

import lombok.Data;
import lombok.experimental.Accessors;

/**
 * 兼容性测试 - 机器信息
 *
 * @author <a href="mailto:472546172@qq.com">张天赐</a>
 */
@Data
@Accessors(chain = true)
public class CompatibilityMachineDTO {
    /**
     * 标识
     */
    private String agentId;
    /**
     * 品牌
     */
    private String brand;
    /**
     * 内存
     */
    private String memory;
    /**
     * 系统类型
     */
    private String osType;
    /**
     * 系统版本
     */
    private String osVersion;
    /**
     * 分辨率
     */
    private String resolution;
    /**
     * 状态
     */
    private Boolean state = true;
    /**
     * 型号
     */
    private String terminalsModel;
    /**
     * 名称
     */
    private String terminalsName;
    /**
     * 序列号
     */
    private String terminalsSerialno;
    /**
     * 类型
     */
    private String terminalsType = "兼容性专用";
}
