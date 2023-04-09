# build
FROM mcr.microsoft.com/dotnet/sdk:latest AS build
WORKDIR /app
RUN git clone https://github.com/AlfredBr/excel-blazor-web-addin.git src
WORKDIR /app/src
RUN dotnet restore
RUN dotnet build -c Release -o /app/out
RUN dotnet publish -c Release -o /app/out

# serve
FROM mcr.microsoft.com/dotnet/aspnet:latest AS serve
WORKDIR /app
COPY --from=build /app/out ./bin
EXPOSE 7061
ENTRYPOINT [ "dotnet", "./bin/excel-blazor-web-addin.dll" ]