FROM mcr.microsoft.com/dotnet/sdk:10.0@sha256:747cbade5bfd918a175ac32f8f1d931edec2c7d0d5397a3de72eaef1ac1f66d2 AS build
WORKDIR /src

COPY global.json Directory.Build.props OfficeAgent.NET.sln ./
COPY src/ ./src/
RUN dotnet restore src/OfficeAgent.Mcp/OfficeAgent.Mcp.csproj --locked-mode

RUN dotnet publish src/OfficeAgent.Mcp/OfficeAgent.Mcp.csproj \
    -c Release -f net10.0 --no-restore -o /app

# The server targets net8.0 and net10.0. .NET 8 support ends 2026-11-10, so the image publishes
# the net10.0 build and runs it on the .NET 10 LTS runtime, built with the SDK global.json pins.
FROM mcr.microsoft.com/dotnet/aspnet:10.0@sha256:8af87eba5e7cdafb1514d7388be287365da97c132509e84dfe252ffd552398ca AS runtime
WORKDIR /app
COPY --from=build /app ./

# The MCP Registry lists this image only when the label names the server in server.json.
LABEL io.modelcontextprotocol.server.name="io.github.ilia-sokolov/officeagent"

ENV OfficeAgent__Transport=http \
    ASPNETCORE_URLS=http://0.0.0.0:8080
EXPOSE 8080

ENTRYPOINT ["dotnet", "OfficeAgent.Mcp.dll"]
