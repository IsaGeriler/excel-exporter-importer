package org.trupt.config;

import org.apache.logging.log4j.Level;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.core.config.Configurator;
import org.apache.logging.log4j.core.config.builder.api.AppenderComponentBuilder;
import org.apache.logging.log4j.core.config.builder.api.ConfigurationBuilder;
import org.apache.logging.log4j.core.config.builder.api.ConfigurationBuilderFactory;

public class Log4j2Config {
    static {
        // Create configuration builder
        ConfigurationBuilder<?> builder = ConfigurationBuilderFactory.newConfigurationBuilder();
        // Console Appender
        AppenderComponentBuilder appenderBuilder = builder.newAppender("Console", "CONSOLE")
                .add(builder.newLayout("PatternLayout")
                .addAttribute("pattern", "%d [%t] %-5level %logger{36} - %msg%n"));
        builder.add(appenderBuilder);
        // Root logger
        builder.add(builder.newRootLogger(Level.DEBUG).add(builder.newAppenderRef("Console")));
        // Initialize the configurator
        Configurator.initialize(builder.build());
    }

    public static Logger getLogger(Class<?> clazz) {
        return LogManager.getLogger(clazz);
    }
}