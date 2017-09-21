package com.talkingtan.httpclient;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;
import com.alibaba.fastjson.JSONPObject;
import org.apache.http.HttpEntity;
import org.apache.http.HttpRequest;
import org.apache.http.HttpResponse;
import org.apache.http.client.HttpClient;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.client.utils.URIBuilder;
import org.apache.http.impl.client.DefaultHttpClient;
import org.apache.http.util.EntityUtils;

import java.net.URI;

/**
 * Http工具类
 * <p>
 * author：Created by ttan on 2017/9/21 0021.
 */
public final class HttpClientUtils {

    public static final String subscriptionKey = "13hc77781f7e4b19b5fcdd72a8df7156";

    public static final String uriBase = "https://westcentralus.api.cognitive.microsoft.com/vision/v1.0/analyze";

    public static JSONObject doPost(String uriBase) throws Exception {
        //httpClient连接
        HttpClient httpClient = new DefaultHttpClient();
        //URL拼接
        URIBuilder uriBuilder = new URIBuilder(uriBase);
        uriBuilder.setParameter("visualFeatures", "Categories,Description,Color");
        uriBuilder.setParameter("language", "en");

        URI uri = uriBuilder.build();

        //Request 请求拼接
        HttpPost request = new HttpPost(uri);
        request.setHeader("Content-Type", "application/json");
        request.setHeader("Ocp-Apim-Subscription-Key", subscriptionKey);

        //Response内容及解析
        HttpResponse response = httpClient.execute(request);
        HttpEntity httpEntity = response.getEntity();

        String entityContent = EntityUtils.toString(httpEntity);
        JSONObject resultObject = JSON.parseObject(entityContent);

        return resultObject;
    }

    public static void main(String[] args) {
        try {
            JSONObject jsonObject = HttpClientUtils.doPost(uriBase);
            System.out.println(jsonObject);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
