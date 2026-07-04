FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

COPY global.json Directory.Build.props OfficeAgent.NET.sln ./
COPY src/ ./src/
RUN dotnet restore src/OfficeAgent.Mcp/OfficeAgent.Mcp.csproj

RUN dotnet publish src/OfficeAgent.Mcp/OfficeAgent.Mcp.csproj \
    -c Release --no-restore -o /app

FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS runtime
WORKDIR /app
COPY --from=build /app ./

ENV OfficeAgent__Transport=http \
    ASPNETCORE_URLS=http://0.0.0.0:8080
EXPOSE 8080

ENTRYPOINT ["dotnet", "OfficeAgent.Mcp.dll"]
