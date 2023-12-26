package com.yupi.springbootinit.utils;

import cn.hutool.core.collection.CollUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.support.ExcelTypeEnum;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.commons.lang3.StringUtils;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * @author lulu
 * @version 1.0
 * @description TODO
 * @date 2023/12/26 19:42
 */
@Slf4j
public class ExcelUtils {


    /**
     * 将Excel文件转换为CSV格式的字符串
     * @param multipartFile 包含Excel文件的MultipartFile对象
     * @return 转换后的CSV格式的字符串，如果文件为空则返回空字符串
     */
    public static String excleToCsv(MultipartFile multipartFile) {
//        File file = null;
//        try {
//            // 从classpath路径下读取名为"网站.xlsx"的文件
//            file = ResourceUtils.getFile("classpath:网站.xlsx");
//        } catch (FileNotFoundException e) {
//            throw new RuntimeException(e);
//        }
        // 读取Excel文件并转换为List<Map<Integer, String>>格式的数据
        List<Map<Integer, String>> list = null;
        try {
            list = EasyExcel.read(multipartFile.getInputStream())
                    .excelType(ExcelTypeEnum.XLSX)
                    .sheet()
                    .headRowNumber(0)
                    .doReadSync();
        } catch (IOException e) {
            log.error("读取Excel文件失败", e);
        }
        if (CollUtil.isEmpty(list)) {
            return "";
        }
        // 转换csv
        StringBuilder stringBuilder = new StringBuilder();
        // 读取表头
        // 将第一行数据转换为LinkedHashMap格式，并赋值给headMap变量
        LinkedHashMap<Integer, String> headMap = (LinkedHashMap) list.get(0);
        // 将headMap中的值提取出来，过滤掉空值，并转换为List<String>格式，赋值给headList变量
        List<String> headList = headMap.values().stream().filter(ObjectUtils::isNotEmpty).collect(Collectors.toList());
        // 将headList中的元素以逗号分隔拼接成字符串，并添加换行符，然后赋值给stringBuilder变量
        stringBuilder.append(StringUtils.join(headList, ",")).append("\n");
        // 读取数据
        for (int i = 1; i < list.size(); i++) {
            // 将当前行数据转换为LinkedHashMap格式，并赋值给dataMap变量
            Map<Integer, String> dataMap = (LinkedHashMap) list.get(i);
            // 将dataMap中的值提取出来，过滤掉空值，并转换为List<String>格式，赋值给dataList变量
            List<String> dataList = dataMap.values().stream().filter(ObjectUtils::isNotEmpty).collect(Collectors.toList());
            // 将dataList中的元素以逗号分隔拼接成字符串，并添加换行符，然后赋值给stringBuilder变量
            stringBuilder.append(StringUtils.join(dataList, ",")).append("\n");
        }
        return stringBuilder.toString();
    }
}
