package com.aici.compatibility;

import java.io.File;
import java.util.List;
import java.io.IOException;
import java.util.ArrayList;
import java.io.FileOutputStream;

import com.aici.compatibility.dto.CompatibilityCommandStatisticDTO;
import com.aici.compatibility.dto.CompatibilityMachineDTO;
import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.Test;

import com.aici.compatibility.dto.CompatibilityCommandStatisticItem;

/**
 * 完整的报告
 *
 * @author <a href="mailto:472546172@qq.com">张天赐</a>
 */
@Slf4j
class PerfectReportTest {
    @Test
    void main() throws IOException {
        long startTime = System.nanoTime();
        List<CompatibilityCommandStatisticItem> items = new ArrayList<>(100);
        for (int i = 0; i < 100; i++) {
            items.add(new CompatibilityCommandStatisticItem(
                new CompatibilityCommandStatisticDTO().setCommandId(i % 2).setTaskId(i / 2L)
                    .setMachine(new CompatibilityMachineDTO().setAgentId("sdfadf" + i)
                        .setBrand("品牌" + i)
                        .setMemory(i + "G").setResolution("100*100" + i)
                        .setOsType(new String[] {"Android", "iOS", "Harmony"}[i % 3])
                        .setOsVersion("1.0." + i)
                        .setTerminalsModel("型号" + i))
                    .setResult(i % 2 == 0)
                    .setMessage(i % 2 == 0 ? "" : "Rsdfafd - "))
                .setStepList(new ArrayList<>(0)));
        }

        log.info("开始生成一份完整的报告");
        File excelFile = new File("perfect/" + System.currentTimeMillis() + ".xlsx");
        if (excelFile.exists() || excelFile.createNewFile()) {
            log.info("excelFile: {}", excelFile.getAbsolutePath());
            try (FileOutputStream excelFileStream = new FileOutputStream(excelFile)) {
                Excel.generate(items, excelFileStream);
                log.info("文件写入完成");
                new Thread(() -> {
                    try {
                        Runtime.getRuntime().exec("open " + excelFile.getAbsolutePath());
                    } catch (IOException e) {log.error("打开文件失败");}
                }).start();
            } catch (Exception e) {log.error("写入Excel文件失败.", e);}
        }
        log.info("耗时:{}", (System.nanoTime() - startTime) / 1_000_000);
    }
}
