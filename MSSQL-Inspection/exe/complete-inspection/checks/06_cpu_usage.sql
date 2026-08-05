WITH rb AS (
  SELECT TOP (1)
    CONVERT(XML, record) AS record_xml
  FROM sys.dm_os_ring_buffers WITH (NOLOCK)
  WHERE ring_buffer_type = 'RING_BUFFER_SCHEDULER_MONITOR'
    AND record LIKE '%<SystemHealth>%'
  ORDER BY timestamp DESC
)
SELECT
  rb.record_xml.value('(Record/SchedulerMonitorEvent/SystemHealth/SystemIdle)[1]', 'int') AS system_idle_percent,
  rb.record_xml.value('(Record/SchedulerMonitorEvent/SystemHealth/ProcessUtilization)[1]', 'int') AS sql_process_cpu_percent,
  (100 - rb.record_xml.value('(Record/SchedulerMonitorEvent/SystemHealth/SystemIdle)[1]', 'int')
   - rb.record_xml.value('(Record/SchedulerMonitorEvent/SystemHealth/ProcessUtilization)[1]', 'int')) AS other_process_cpu_percent
FROM rb
OPTION (RECOMPILE);
