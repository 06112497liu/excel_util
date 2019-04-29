package com.tellyes.core.utils;

import com.tellyes.core.constants.Constants;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintStream;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.Optional;
import java.util.function.BiFunction;
import java.util.function.Predicate;

import static com.tellyes.core.constants.EncodingTypeEnum.ISO_8859_1;
import static com.tellyes.core.constants.EncodingTypeEnum.UTF_8;

/**
 * 下载工具类
 * @author xiehai
 * @date 2018/04/26 17:22
 * @Copyright(c) tellyes tech. inc. co.,ltd
 */
public interface DownloadUtil {
    /**
     * 格式化下载文件名函数
     */
    BiFunction<String, String, String> FORMAT_FILE_NAME = (name, suffix) ->
        String.format("attachment; filename=\"%s.%s\"", name, suffix);
    /**
     * 判断是否是ie内核浏览器断言
     */
    Predicate<String> IS_IE = userAgent ->
        // 是否是ie浏览器
        userAgent.contains("MSIE")
            // 是否是edge浏览器
            || userAgent.contains("EDGE")
            // 是否是ie内核浏览器
            || userAgent.contains("TRIDENT");
    /**
     * http文件下载
     * @param request  http request
     * @param response http response
     * @param name     文件名
     * @param inStream 文件流
     * @param suffix   文件后缀
     * @throws Exception IOException
     */
    static void download(HttpServletRequest request, HttpServletResponse response, String name, InputStream inStream,
                         String suffix)
        throws Exception {
        // 设置下载文件名
        String newFileName =
            Optional.of(request.getHeader(HttpHeaders.USER_AGENT).toUpperCase())
                .filter(IS_IE)
                .map(t -> {
                    try {
                        return URLEncoder.encode(name, UTF_8.getType());
                    } catch (UnsupportedEncodingException e) {
                        return Constants.EMPTY;
                    }
                }).orElse(new String(name.getBytes(UTF_8.getType()), ISO_8859_1.getType()));

        response.setContentType(MediaType.APPLICATION_OCTET_STREAM_VALUE);
        response.setHeader(HttpHeaders.CONTENT_DISPOSITION, FORMAT_FILE_NAME.apply(newFileName, suffix));

        byte[] b = new byte[100];
        try (OutputStream outStream = response.getOutputStream();
             PrintStream out = new PrintStream(outStream, true, UTF_8.getType())) {
            int len;
            while ((len = inStream.read(b)) > 0) {
                out.write(b, 0, len);
                out.flush();
            }
        } finally {
            inStream.close();
        }
    }
}
