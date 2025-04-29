// program.cs tambahin service buat csv excel


// mapper.cs
CreateMap<CsvItemDTO, ItemDTO>()
			.ForMember(dest => dest.VolumeProperty, opt => opt.MapFrom(src => src.VolumeProperty ?? "0.0,0.0,0.0"))
			.ForMember(dest => dest.PicturePath, opt => opt.MapFrom(src => src.PicturePath ?? "default_item.png"))
			.ForMember(dest => dest.ExpiryReminder, opt => opt.MapFrom(src => 30));


// csvitemdto.cs
using CsvHelper.Configuration.Attributes;

namespace InventoryManagementSystem.Models.ModelsDTO;
public class CsvItemDTO
{
	[Name("Item Name")]
	public string ItemName { get; set; }
	[Name("Description")]
	public string? Description { get; set; }
	[Name("Expire Date")]
	public DateTime? ExpireDate { get; set; }
	[Name("Volume Property")]
	public string? VolumeProperty { get; set; }
	[Name("Picture Path")]
	public string? PicturePath { get; set; }
	[Name("Availability")]
	public bool Availability { get; set; }	
	[Name("Quantity")]
	public int Quantity { get; set; }
	[Name("Quantity Low")]
	public int QuantityLow { get; set; }
	[Name("Type ID")]
	public int TypeId { get; set; }
	[Name("Category Name")]
	public string CategoryName { get; set; }
	[Name("Subcategory Name")]
	public string SubCategoryName { get; set; }
	[Name("City Name")]
	public string CityName { get; set; }
}

// igetentity and getentityservice
namespace InventoryManagementSystem.Areas.Admin.Services.IServices;
public interface IGetEntityIdService
{
    Task<Guid> GetCategoryId(string categoryName);
    Task<Guid> GetSubCategoryId(string subCategoryName, Guid categoryId);
    Task<Guid> GetCityId(string cityName);
}

namespace InventoryManagementSystem.Areas.Admin.Services;

public class GetEntityIdService : IGetEntityIdService
{
    private readonly ICategoryRepository _category;
    private readonly ISubCategoryRepository _subCategory;
    private readonly ICityRepository _city;
    private readonly ILogger<GetEntityIdService> _logger;

    public GetEntityIdService(
        ICategoryRepository category,
        ISubCategoryRepository subCategory,
        ICityRepository city,
        ILogger<GetEntityIdService> logger)
    {
        _category = category;
        _subCategory = subCategory;
        _city = city;
        _logger = logger;
    }

    public async Task<Guid> GetCategoryId(string categoryName)
    {
        var category = (await _category.GetAll()).FirstOrDefault(c => 
            c.CategoryName.Equals(categoryName, StringComparison.OrdinalIgnoreCase));

        if (category == null)
        {
            _logger.LogError($"Category '{categoryName}' not found in the system");
            throw new Exception($"Category '{categoryName}' not found in the system");
        }

        return category.CategoryId;
    }

    public async Task<Guid> GetSubCategoryId(string subCategoryName, Guid categoryId)
    {
        var subCategory = (await _subCategory.GetAll()).FirstOrDefault(s => 
            s.SubCategoryName.Equals(subCategoryName, StringComparison.OrdinalIgnoreCase) && 
            s.CategoryId == categoryId);

        if (subCategory == null)
        {
            _logger.LogError($"Subcategory '{subCategoryName}' not found for selected category");
            throw new Exception($"Subcategory '{subCategoryName}' not found for selected category");
        }

        return subCategory.SubCategoryId;
    }

    public async Task<Guid> GetCityId(string cityName)
    {
        var city = (await _city.GetAll()).FirstOrDefault(c => 
            c.CityName.Equals(cityName, StringComparison.OrdinalIgnoreCase));

        if (city == null)
        {
            _logger.LogError($"City '{cityName}' not found in the system");
            throw new Exception($"City '{cityName}' not found in the system");
        }

        return city.CityId;
    }
}


// ifileimport and fileimportservice
namespace InventoryManagementSystem.Areas.Admin.Services.IServices;
public interface IFileImportService
{
    Task<string> SaveFileAsync(IFormFile file, string fileExtension);
    Task<bool> AddItemToDatabaseAsync(ItemDTO itemDTO);
    Task<(bool Success, string Message, List<ItemDTO> Items)> ProcessFileAsync(IFormFile file);
    Task<bool> ImportItemsToDatabase(List<ItemDTO> items);
}

public abstract class FileImportService : IFileImportService
{
    protected readonly IItemRepository _item;
    protected readonly IItemViewRepository _itemView;
    protected readonly IGetEntityIdService _getEntityIdService;
    protected readonly ItemViewAdminService _itemViewService;
    protected readonly IMapper _mapper;
    protected readonly ILogger _logger;

    public FileImportService(
        IItemRepository item,
        IItemViewRepository itemView,
        IGetEntityIdService getEntityIdService,
        ItemViewAdminService itemViewService,
        IMapper mapper,
        ILogger logger)
    {
        _item = item;
        _itemView = itemView;
        _getEntityIdService = getEntityIdService;
        _itemViewService = itemViewService;
        _mapper = mapper;
        _logger = logger;
    }

    public async Task<string> SaveFileAsync(IFormFile file, string fileExtension)
    {
        if (file == null || file.Length == 0)
        {
            return null;
        }

        if (!file.FileName.EndsWith(fileExtension))
        {
            throw new Exception($"Only {fileExtension} files are allowed");
        }

        var uploadFolder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Uploads");

        if (!Directory.Exists(uploadFolder))
        {
            Directory.CreateDirectory(uploadFolder);
        }

        var filePath = Path.Combine(uploadFolder, file.FileName);

        try
        {
            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }
            return filePath;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to save file: {Message}", ex.Message);
            return null;
        }
    }

    public async Task<bool> AddItemToDatabaseAsync(ItemDTO itemDTO)
    {
        try
        {
            // Get IDs for category, subcategory, and city
            var categoryId = await _entityLookupService.GetCategoryId(itemDTO.CategoryName);
            var subCategoryId = await _entityLookupService.GetSubCategoryId(itemDTO.SubCategoryName, categoryId);
            var cityId = await _entityLookupService.GetCityId(itemDTO.CityName);

            // Generate item code
            var itemCode = await _itemViewService.GenerateItemCode(categoryId, subCategoryId);

            // Map ItemDTO to Item
            var item = _mapper.Map<Item>(itemDTO);
            item.CategoryId = categoryId;
            item.SubCategoryId = subCategoryId;
            item.CityId = cityId;
            item.ItemCode = itemCode;
            item.CreateAt = DateTime.Now;
            item.IsDeleted = false;
            item.ExpiryReminder = itemDTO.ExpiryReminder ?? 30; // Default value

            // Add item to the database
            await _item.AddAsync(item);
            await _item.SaveAsync();

            // Add the item reference to the Item Views Table
            var itemView = new ItemView
            {
                ItemId = item.ItemId,
                IsDeleted = false
            };
            await _itemView.AddAsync(itemView);
            await _itemView.SaveAsync();

            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error adding item to database: {Message}", ex.Message);
            return false;
        }
    }

    public abstract Task<(bool Success, string Message, List<ItemDTO> Items)> ProcessFileAsync(IFormFile file);
    public abstract Task<bool> ImportItemsToDatabase(List<ItemDTO> items);
}


// ICsv and csvservice
public interface ICsvService
{
    Task<(bool Success, string Message, List<ItemDTO> Items)> ProcessCsvFileAsync(IFormFile file);
    Task<bool> ImportItemsToDatabase(List<ItemDTO> items);
}

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
using InventoryManagementSystem.Areas.Admin.Services;
using AutoMapper;
using System;

namespace InventoryManagementSystem.Areas.Admin.Services;

public class CsvService : ICsvService
{
    private readonly IItemRepository _item;
    private readonly IItemViewRepository _itemView;
    private readonly ICategoryRepository _category;
    private readonly ISubCategoryRepository _subCategory;
    private readonly ICityRepository _city;
    private readonly ItemViewAdminService _itemViewService;
    private readonly IMapper _mapper;
    private readonly ILogger<CsvService> _logger;

    public CsvService(
        IItemRepository item,
        IItemViewRepository itemView,
        IEntityLookupService entityLookupService,
        ItemViewAdminService itemViewService,
        IMapper mapper,
        ILogger<CsvService> logger) 
        : base(item, itemView, entityLookupService, itemViewService, mapper, logger)
    {
        _item = item;
        _itemView = itemView;
        _category = entityLookupService.CategoryRepository;
        _subCategory = entityLookupService.SubCategoryRepository;
        _city = entityLookupService.CityRepository;
        _itemViewService = itemViewService;
        _mapper = mapper;
        _logger = logger;
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

        if (!file.FileName.EndsWith(".csv"))
        {
            return (false, "Only CSV files are allowed", null);
        }

        try
        {
            // First, read the CsvItemDTO objects with CsvHelper
            var csvItems = new List<CsvItemDTO>();
            using (var reader = new StreamReader(file.OpenReadStream()))
            using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HeaderValidated = null,
                MissingFieldFound = null
            }))
            {
                csvItems = csv.GetRecords<CsvItemDTO>().ToList();
            }

            if (csvItems.Count == 0)
            {
                return (false, "No records found in the CSV file", null);
            }

            // Then map to our unified ItemDTO
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



// iexcel and excelservice
namespace InventoryManagementSystem.Areas.Admin.Services.IServices;

public interface ICsvService
{
    Task<(bool Success, string Message, List<ItemDTO> Items)> ProcessCsvFileAsync(IFormFile file);
    Task<bool> ImportItemsToDatabase(List<ItemDTO> items);
}

di cld excelservice


//controller jgn lu-pa inject
[HttpPost]
	public async Task<IActionResult> UploadFile(IFormFile file, string fileType)
	{
		if (file == null || file.Length == 0)
		{
			TempData["Error"] = "Please select a file to upload";
			return RedirectToAction(nameof(Index));
		}

		try
		{
			bool success;
			string message;
			List<ItemDTO> items;
			bool importResult;

			// Choose service based on file type
			if (fileType.Equals("csv", StringComparison.OrdinalIgnoreCase))
			{
				(success, message, items) = await _csvService.ProcessFileAsync(file);
				if (!success)
				{
					TempData["Error"] = message;
					return RedirectToAction(nameof(Index));
				}
				importResult = await _csvService.ImportItemsToDatabase(items);
			}
			else if (fileType.Equals("excel", StringComparison.OrdinalIgnoreCase))
			{
				(success, message, items) = await _excelService.ProcessFileAsync(file);
				if (!success)
				{
					TempData["Error"] = message;
					return RedirectToAction(nameof(Index));
				}
				importResult = await _excelService.ImportItemsToDatabase(items);
			}
			else
			{
				TempData["Error"] = "Unsupported file type";
				return RedirectToAction(nameof(Index));
			}

			if (importResult)
			{
				TempData["Success"] = $"Successfully imported {items.Count} items from {fileType.ToUpper()}";
				_logger.LogInformation("Successfully imported {count} items via {fileType} upload", items.Count, fileType);
			}
			else
			{
				TempData["Error"] = "Error occurred while importing data to database";
			}

		}
		catch (Exception ex)
		{
			_logger.LogError(ex, "Error processing {fileType} upload: {Message}", fileType, ex.Message);
			TempData["Error"] = $"Error processing {fileType.ToUpper()}: {ex.Message}";
		}

		return RedirectToAction(nameof(Index));
	}







































