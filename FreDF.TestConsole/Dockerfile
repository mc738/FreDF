FROM mcr.microsoft.com/dotnet/runtime:5.0 AS base
WORKDIR /app

FROM mcr.microsoft.com/dotnet/sdk:5.0 AS build
WORKDIR /src
COPY ["FreDF.TestConsole/FreDF.TestConsole.fsproj", "FreDF.TestConsole/"]
RUN dotnet restore "FreDF.TestConsole/FreDF.TestConsole.fsproj"
COPY . .
WORKDIR "/src/FreDF.TestConsole"
RUN dotnet build "FreDF.TestConsole.fsproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "FreDF.TestConsole.fsproj" -c Release -o /app/publish

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "FreDF.TestConsole.dll"]
