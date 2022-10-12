package com.aici.compatibility.dto;

import java.util.List;

import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.experimental.Accessors;

/**
 * 兼容性测试 命令执行的固化数据
 *
 * @author <a href="mailto:472546172@qq.com">张天赐</a>
 */
@Data
@Accessors(chain = true)
@EqualsAndHashCode(callSuper = true)
public class CompatibilityCommandStatisticItem extends CompatibilityCommandStatisticDTO {
    public CompatibilityCommandStatisticItem() {}

    public CompatibilityCommandStatisticItem(CompatibilityCommandStatisticDTO dto) {
        this.setCommandId(dto.getCommandId())
            .setData(dto.getData())
            .setMachine(dto.getMachine())
            .setCaseId(dto.getCaseId())
            .setFinishTime(dto.getFinishTime())
            .setInQueueTime(dto.getInQueueTime())
            .setResult(dto.getResult())
            .setMessage(dto.getMessage())
            .setTaskId(dto.getTaskId())
        ;
    }

    /**
     * 关联的测试步骤
     */
    private List<CompatibilityStepDTO> stepList;
}
