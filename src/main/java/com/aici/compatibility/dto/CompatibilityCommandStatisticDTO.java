package com.aici.compatibility.dto;

import java.util.Map;
import java.util.HashMap;

import lombok.Data;
import lombok.experimental.Accessors;

/**
 * 兼容性测试 命令执行的固化数据
 *
 * @author <a href="mailto:472546172@qq.com">张天赐</a>
 */
@Data
@Accessors(chain = true)
public class CompatibilityCommandStatisticDTO {
    /**
     * 执行结果
     */
    private Boolean result;
    /**
     * 执行结果的描述
     */
    private String message;
    /**
     * 用例主键
     */
    private Long caseId;
    /**
     * 任务主键
     */
    private Long taskId;
    /**
     * 命令序号
     */
    private Integer commandId;
    /**
     * 机器标识
     */
    private String machineId;
    /**
     * 入队时间
     */
    private Long inQueueTime;
    /**
     * 结束时间
     */
    private Long finishTime;
    /**
     * 统计数据
     */
    private Map<String, Object> data = new HashMap<>(4);
    /**
     * 机器信息
     */
    private CompatibilityMachineDTO machine;
}
