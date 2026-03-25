
from fastmcp import Client
from fastmcp.client.transports import StdioTransport

import config_huggingface as config

class MCPManager:
    def __init__(self):
        self.client = None
        self._client_cm = None

    async def connect(self):
        # Use explicitly defined Python to avoid 32/64 bit ODBC mismatches
        transport = StdioTransport(
            command=config.MCP_COMMAND,
            args=[config.MCP_SCRIPT]
        )
        self.client = Client(transport)
        self._client_cm = self.client
        await self._client_cm.__aenter__()

    async def close(self):
        if self._client_cm:
            await self._client_cm.__aexit__(None, None, None)

    async def list_tools(self):
        tools = await self.client.list_tools()
        return tools  # list of Tool objects, same as before

    async def call(self, tool, args):
        result = await self.client.call_tool(tool, args)
        # FastMCP may return the result directly or wrapped
        # Check if it's a string or needs extraction
        if isinstance(result, str):
            return result
        # If it has content attribute, extract it
        if hasattr(result, 'content') and result.content:
            if isinstance(result.content, list) and len(result.content) > 0:
                return result.content[0].text if hasattr(result.content[0], 'text') else str(result.content[0])
            return str(result.content)
        return str(result)
    