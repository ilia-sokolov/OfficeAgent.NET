FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

COPY global.json Directory.Build.props OfficeAgent.NET.sln ./
COPY src/ ./src/
RUN dotnet restore src/OfficeAgent.Mcp/OfficeAgent.Mcp.csproj

RUN dotnet publish src/OfficeAgent.Mcp/OfficeAgent.Mcp.csproj \
    -c Release --no-restore -o /app

# The server targets net8.0 with RollForward=LatestMajor. .NET 8 support ends 2026-11-10, so
# the image runs it on the .NET 10 LTS runtime; the build stage keeps the pinned .NET 8 SDK.
FROM mcr.microsoft.com/dotnet/aspnet:10.0 AS runtime
WORKDIR /app
COPY --from=build /app ./

ENV OfficeAgent__Transport=http \
    ASPNETCORE_URLS=http://0.0.0.0:8080
EXPOSE 8080

ENTRYPOINT ["dotnet", "OfficeAgent.Mcp.dll"]
