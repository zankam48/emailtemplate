using CsvHelper;
using CsvHelper.Configuration;
using System.Globalization;
using InventoryManagementSystem.Models.ModelsDTO;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using InventoryManagementSystem.Persistence.Repository.IRepository;
using InventoryManagementSystem.Models.Models;
using InventoryManagementSystem.Areas.Admin.Services.IServices;
using AutoMapper;
using System;

namespace InventoryManagementSystem.Areas.Admin.Services;

public class CsvService : FileImportService, ICsvService
{
    public CsvService(
        IItemRepository item,
        IItemViewRepository itemView,
        IGetEntityIdService getEntityIdService,
        ItemViewAdminService itemViewService,
        IMapper mapper,
        ILogger<CsvService> logger) 
        : base(item, itemView, getEntityIdService, itemViewService, mapper, logger)
    {
    }

    public override async Task<(bool Success, string Message, List<ItemDTO> Items)> ProcessFileAsync(IFormFile file)
    {
        return await ProcessCsvFileAsync(file);
    }

    public async Task<(bool Success, string Message, List<ItemDTO> Items)> ProcessCsvFileAsync(IFormFile file)
    {
        if (file == null || file.Length == 0)
        {
            return (false, "File is empty", null);
        }

        if (!file.FileName.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
        {
            return (false, "Only CSV files are allowed", null);
        }

        try
        {
            var csvItems = new List<ItemDTO>();
            using (var reader = new StreamReader(file.OpenReadStream()))
            using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HeaderValidated = null,
                MissingFieldFound = null
            }))
            {
                csv.Context.RegisterClassMap<CsvItemDTOMap>();
                csvItems = csv.GetRecords<ItemDTO>().ToList();
            }

            if (csvItems.Count == 0)
            {
                return (false, "No records found in the CSV file", null);
            }

            var items = _mapper.Map<List<ItemDTO>>(csvItems);

            return (true, $"Successfully processed {items.Count} records", items);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error processing CSV file: {Message}", ex.Message);
            return (false, $"Error processing CSV file: {ex.Message}", null);
        }
    }

    public override async Task<bool> ImportItemsToDatabase(List<ItemDTO> items)
    {
        try
        {
            foreach (var item in items)
            {
                await AddItemToDatabaseAsync(item);
            }
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error importing items to database: {Message}", ex.Message);
            return false;
        }
    }
}
