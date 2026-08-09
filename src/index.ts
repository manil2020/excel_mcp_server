#!/usr/bin/env node
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { createServer } from './server.js';
import { log } from './util/logger.js';

async function main(): Promise<void> {
  const { server, manager } = createServer();
  const transport = new StdioServerTransport();

  const shutdown = async (signal: NodeJS.Signals): Promise<void> => {
    log.info('shutting down', { signal });
    try {
      manager.closeAll();
    } catch (e) {
      log.warn('closeAll failed during shutdown', { error: (e as Error).message });
    }
    try {
      await server.close();
    } catch (e) {
      log.warn('server.close failed during shutdown', { error: (e as Error).message });
    }
    process.exit(0);
  };

  process.on('SIGINT', () => void shutdown('SIGINT'));
  process.on('SIGTERM', () => void shutdown('SIGTERM'));

  await server.connect(transport);
  log.info('MCP server listening on stdio');
}

main().catch((err) => {
  log.error('fatal error in main', { error: (err as Error).message, stack: (err as Error).stack });
  process.exit(1);
});
