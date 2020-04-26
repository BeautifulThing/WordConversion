package com.kland.common.config;

import org.springframework.boot.web.context.WebServerInitializedEvent;
import org.springframework.context.ApplicationListener;
import org.springframework.context.annotation.Configuration;

import java.net.InetAddress;
import java.net.UnknownHostException;

@Configuration
public class ServerConfig implements ApplicationListener<WebServerInitializedEvent> {
    private Integer serverPort;
    @Override
    public void onApplicationEvent(WebServerInitializedEvent serverInitializedEvent) {
        this.serverPort = serverInitializedEvent.getWebServer().getPort();
    }
    public String getUrl() {
        InetAddress address = null;
        try {
            address = InetAddress.getLocalHost();
        } catch (UnknownHostException e) {
            e.printStackTrace();
        }
        return "http://"+address.getHostAddress() +":"+this.serverPort;
    }
}