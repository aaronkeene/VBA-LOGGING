# VBA-LOGGING
Provides a standardized facility for handling log messages in a VBA application.  This library was heavily influenced by the Python logging library and there is a lot of similar functionality.

# Basic Logging

## Logging Levels

Logging functions are named after the level or severity of the events they are used to track. The standard levels and their applicability are described below (in increasing order of severity):

| Level    | When to Use |
| :---     | :--- | 
| CRITICAL | Indicates a serious error where the application may fail to continue running  |
| ERROR    | Indicates an application error where some function failed  |
| WARNING  | Indicates  something unexpected occured and/or warns of furture errors (e.g. 'low disk spave'). Typically software continuies to work as expected |
| INFO     | Provides confirmation things are working as expected |
| DEBUG    | Provides detailed information to support troubleshooting / diagnosis |

Default level is `WARNING`, which means events of this level and above will be tracked unless package is configured to do otherwise.

Tracked events (e.g. logs) can be handled many different ways, however the simplest (and default way) is to print to the console (VBA immediate window).

# Console Example

Impor the logging modules into the VBA project. Note most of the class objects have `VB_PredeclaredId` set to `True` so you get a default instance available. To log to the immediate window, add the following lines to your procedure

```
Logging.LogWarning ("Look out!") 'Prints to immediate window
Logging.LogInfo ("Told you so.") 'Will not print to immediate window
```

If you type these lines into a script and run, you’ll see:

```
WARNING - RootLogger - 2023-12-31 17:40:09 - Look out!
```


# Advanced Logging

The logging library takes a modular approach and offers several categories of components: loggers, handlers, filters, and formatters.
- Loggers expose the interface that application code directly uses.
- Handlers send the log records (created by loggers) to the appropriate destination.
- Filters provide a finer grained facility for determining which log records to output.
- Formatters specify the layout of log records in the final output.

## Logging Flow

## Loggers

## Handlers

## Filters
Not implemented yet
