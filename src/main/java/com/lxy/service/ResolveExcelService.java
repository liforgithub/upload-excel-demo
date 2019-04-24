package com.lxy.service;

import org.springframework.web.multipart.MultipartFile;

/**
 * <p>
 * 类说明
 * </p>
 *
 * @author lxy
 * @since 2019/4/24
 */
public interface ResolveExcelService {


    /**
     * 解析Excel
     *
     * @param file 文件
     * @return 得到的结果
     * @throws Exception 业务异常统一处理
     */
    void resolveExcel(MultipartFile file) throws Exception;
}
