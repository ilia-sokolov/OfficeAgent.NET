FROM mcr.microsoft.com/dotnet/sdk:10.0@sha256:de4fe90a98845d75910b9769662465c84de1a934665c431279d1382b2bddc02b AS build
WORKDIR /src

COPY global.json Directory.Build.props OfficeAgent.NET.sln ./
COPY src/ ./src/
RUN dotnet restore src/OfficeAgent.Mcp/OfficeAgent.Mcp.csproj

RUN dotnet publish src/OfficeAgent.Mcp/OfficeAgent.Mcp.csproj \
    -c Release -f net10.0 --no-restore -o /app

# The server targets net8.0 and net10.0. .NET 8 support ends 2026-11-10, so the image publishes
# the net10.0 build and runs it on the .NET 10 LTS runtime, built with the SDK global.json pins.
FROM mcr.microsoft.com/dotnet/aspnet:10.0@sha256:de0ee358cdd672f7017a003bed21e9ac9ed48a27c922a59b18175b6e82fa7fee AS runtime
WORKDIR /app
COPY --from=build /app ./

# The MCP Registry lists this image only when the label names the server in server.json.
LABEL io.modelcontextprotocol.server.name="io.github.ilia-sokolov/officeagent"

ENV OfficeAgent__Transport=http \
    ASPNETCORE_URLS=http://0.0.0.0:8080
EXPOSE 8080

ENTRYPOINT ["dotnet", "OfficeAgent.Mcp.dll"]
