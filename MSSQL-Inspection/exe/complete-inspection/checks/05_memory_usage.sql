SELECT
  opm.physical_memory_in_use_kb / 1024.0 AS sql_memory_in_use_mb,
  opm.locked_page_allocations_kb / 1024.0 AS locked_pages_mb,
  opm.large_page_allocations_kb / 1024.0 AS large_pages_mb,
  osm.total_physical_memory_kb / 1024.0 AS host_total_memory_mb,
  osm.available_physical_memory_kb / 1024.0 AS host_available_memory_mb,
  CASE
    WHEN osm.total_physical_memory_kb > 0
      THEN (opm.physical_memory_in_use_kb * 100.0) / osm.total_physical_memory_kb
    ELSE NULL
  END AS sql_memory_percent_of_host
FROM sys.dm_os_process_memory AS opm
CROSS JOIN sys.dm_os_sys_memory AS osm
OPTION (RECOMPILE);
