package com.yupi.springbootinit.model.dto.chart;

import com.sun.xml.internal.bind.v2.TODO;
import lombok.Data;

import java.io.Serializable;

/**
 * 文件上传请求
 *
 */
@Data
public class GenChartByAiRequest implements Serializable {

    /**
     * 名称
     */
    private String name;

    /**
     * 分析目标
     */
    private String goal;


    /**
     * 图表类型
     */
    private String chartType;

    private static final long serialVersionUID = 1L;
}
