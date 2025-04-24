# Use the .NET 6.0 SDK image to build the project
FROM mcr.microsoft.com/dotnet/sdk:6.0 AS build
WORKDIR /app

# Copy all files and restore dependencies
COPY . ./
RUN dotnet restore

# Build and publish the app in Release mode
RUN dotnet publish -c Release -o out

# Use the .NET 6.0 ASP.NET runtime image to run the app
FROM mcr.microsoft.com/dotnet/aspnet:6.0
WORKDIR /app

# Copy the published files from the build stage
COPY --from=build /app/out .

# Ensure the Data folder (with ALPON.xlsx) is copied into the container
COPY Data /app/Data

# Set the entry point (replace with your DLL name if different)
CMD ["dotnet", "TelegramExcelBot.dll"]
