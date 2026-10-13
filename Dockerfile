FROM mcr.microsoft.com/dotnet/sdk:10.0@sha256:0d99ebf1ccad15037b5c98c9fe8979d6b8dbac23812feae2277a0b277b5c5ffa AS build
WORKDIR /src

COPY global.json Directory.Build.props OfficeAgent.NET.sln ./
COPY src/ ./src/
RUN dotnet restore src/OfficeAgent.Mcp/OfficeAgent.Mcp.csproj

RUN dotnet publish src/OfficeAgent.Mcp/OfficeAgent.Mcp.csproj \
    -c Release -f net10.0 --no-restore -o /app

# The server targets net8.0 and net10.0. .NET 8 support ends 2026-11-10, so the image publishes
# the net10.0 build and runs it on the .NET 10 LTS runtime, built with the SDK global.json pins.
FROM mcr.microsoft.com/dotnet/aspnet:10.0@sha256:d5633778c754e89a386882eddf8dbee6942447a8a4b8267c19cf1e1158d0b0ac AS runtime
WORKDIR /app
COPY --from=build /app ./

# The MCP Registry lists this image only when the label names the server in server.json.
LABEL io.modelcontextprotocol.server.name="io.github.ilia-sokolov/officeagent"

ENV OfficeAgent__Transport=http \
    ASPNETCORE_URLS=http://0.0.0.0:8080
EXPOSE 8080

ENTRYPOINT ["dotnet", "OfficeAgent.Mcp.dll"]
