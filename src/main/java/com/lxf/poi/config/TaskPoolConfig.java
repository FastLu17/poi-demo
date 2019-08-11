package com.lxf.poi.config;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.scheduling.concurrent.ThreadPoolTaskExecutor;

import java.util.concurrent.Executor;
import java.util.concurrent.ThreadPoolExecutor;

/**
 * 使用 自定义的线程池 完成@Async异步任务、
 * <p>
 * 需要在@Async(value="taskExecutor")注解中指定value的值(绑定自定义的线程池)。
 * <p>
 * 配置此类后，@Async不绑定也是使用此线程池执行。
 *
 * @author: 小66
 * @Description:
 * @create: 2019-08-11 15:43
 **/
@Configuration
public class TaskPoolConfig {

    @Bean("taskExecutor")
    public Executor taskExecutor() {
        ThreadPoolTaskExecutor executor = new ThreadPoolTaskExecutor();
        executor.setCorePoolSize(10);
        executor.setMaxPoolSize(20);
        executor.setQueueCapacity(200);//设置队列容量
        executor.setKeepAliveSeconds(60);
        executor.setThreadNamePrefix("taskExecutor-");//设置线程的名称前缀
        executor.setRejectedExecutionHandler(new ThreadPoolExecutor.CallerRunsPolicy());
        return executor;
    }
}
