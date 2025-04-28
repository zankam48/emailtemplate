//csv service
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

public interface ICsvService
{
    Task<(bool Success, string Message, List<CsvItemDTO> Items)> ProcessCsvFileAsync(IFormFile file);
    Task<bool> ImportItemsToDatabase(List<CsvItemDTO> items);
}

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
        ICategoryRepository category,
        ISubCategoryRepository subCategory,
        ICityRepository city,
        ItemViewAdminService itemViewService,
        IMapper mapper,
        ILogger<CsvService> logger)
    {
        _item = item;
        _itemView = itemView;
        _category = category;
        _subCategory = subCategory;
        _city = city;
        _itemViewService = itemViewService;
        _mapper = mapper;
        _logger = logger;
    }

    public async Task<(bool Success, string Message, List<CsvItemDTO> Items)> ProcessCsvFileAsync(IFormFile file)
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
            var items = new List<CsvItemDTO>();

            using (var reader = new StreamReader(file.OpenReadStream()))
            using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HeaderValidated = null,
                MissingFieldFound = null
            }))
            {
                items = csv.GetRecords<CsvItemDTO>().ToList();
            }

            if (items.Count == 0)
            {
                return (false, "No records found in the CSV file", null);
            }

            return (true, $"Successfully processed {items.Count} records", items);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error processing CSV file: {Message}", ex.Message);
            return (false, $"Error processing CSV file: {ex.Message}", null);
        }
    }

    public async Task<bool> ImportItemsToDatabase(List<CsvItemDTO> items)
    {
        try
        {
            foreach (var csvItem in items)
            {
                await AddItemToDatabaseAsync(csvItem);
            }
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error importing items to database: {Message}", ex.Message);
            return false;
        }
    }

    private async Task AddItemToDatabaseAsync(CsvItemDTO csvItemDTO)
    {
        // Get or create category
        var categoryId = await GetCategoryId(csvItemDTO.CategoryName);
        
        // Get or create subcategory
        var subCategoryId = await GetSubCategoryId(csvItemDTO.SubCategoryName, categoryId);
        
        // Get or create city
        var cityId = await GetCityId(csvItemDTO.CityName);

        // Generate item code based on category and subcategory
        var itemCode = await _itemViewService.GenerateItemCode(categoryId, subCategoryId);

        // Map CSV DTO to Item model
        var item = new Item
        {
            ItemName = csvItemDTO.ItemName,
            Description = csvItemDTO.Description,
            ExpireDate = csvItemDTO.ExpireDate,
            Quantity = csvItemDTO.Quantity,
            QuantityLow = csvItemDTO.QuantityLow,
            Availability = csvItemDTO.Availability,
            PicturePath = csvItemDTO.PicturePath ?? "default_item.png", // Default picture path
            VolumeProperty = "0.0,0.0,0.0", // Default volume property
            TypeId = csvItemDTO.TypeId,
            CategoryId = categoryId,
            SubCategoryId = subCategoryId,
            CityId = cityId,
            ItemCode = itemCode,
            CreateAt = DateTime.Now,
            IsDeleted = false,
            ExpiryReminder = 30 // Default expiry reminder in days
        };

        // Add item to database
        await _item.AddAsync(item);
        await _item.SaveAsync();

        // Create ItemView record
        var itemView = new ItemView
        {
            ItemId = item.ItemId,
            IsDeleted = false
        };

        await _itemView.AddAsync(itemView);
        await _itemView.SaveAsync();
    }

    private async Task<Guid> GetCategoryId(string categoryName)
    {
        var category = (await _category.GetAll()).FirstOrDefault(c => 
            c.CategoryName.Equals(categoryName, StringComparison.OrdinalIgnoreCase));

        if (category == null)
        {
            throw new Exception($"Category '{categoryName}' not found in the system");
        }

        return category.CategoryId;
    }

    private async Task<Guid> GetSubCategoryId(string subCategoryName, Guid categoryId)
    {
        var subCategory = (await _subCategory.GetAll()).FirstOrDefault(s => 
            s.SubCategoryName.Equals(subCategoryName, StringComparison.OrdinalIgnoreCase) && 
            s.CategoryId == categoryId);

        if (subCategory == null)
        {
            throw new Exception($"Subcategory '{subCategoryName}' not found for selected category");
        }

        return subCategory.SubCategoryId;
    }

    private async Task<Guid> GetCityId(string cityName)
    {
        var city = (await _city.GetAll()).FirstOrDefault(c => 
            c.CityName.Equals(cityName, StringComparison.OrdinalIgnoreCase));

        if (city == null)
        {
            throw new Exception($"City '{cityName}' not found in the system");
        }

        return city.CityId;
    }
}


//excelservice
using InventoryManagementSystem.Persistence.Repository.IRepository;

namespace InventoryManagementSystem.Areas.Admin.Services;

public class ExcelService
{
    private readonly ICategoryRepository _category;
    private readonly ISubCategoryRepository _subCategory; 
    private readonly ICityRepository _city;
    
    public ExcelService(ICategoryRepository category,
    ISubCategoryRepository subCategory,
    ICityRepository city
    ) 
    {
        _category = category;
        _subCategory = subCategory;
        _city = city;
    }

    private async Task<Guid?> GetCategoryCode(string categoryId)
    {
        return (await _category.GetAll()).FirstOrDefault(c => c.CategoryName == categoryId)?.CategoryId;
    }

    private async Task<string> GetSubCategoryCode(Guid subCategoryId)
    {
        return (await _subCategory.GetAll()).FirstOrDefault(s => s.SubCategoryId == subCategoryId)?.SubCategoryCode;
    }

    [HttpPost]
	public async Task<IActionResult> Upload(ItemViewFilter itemViewFilter)
	{
		if (!ModelState.IsValid)
		{
			return View(itemViewFilter);
		}
		if (itemViewFilter.File == null || itemViewFilter.File.Length == 0)
		{
			ModelState.AddModelError("", "File Not Found.");
			return BadRequest("File Not Found.");
		}

		var filePath = await SaveFileAsync(itemViewFilter.File);

		if (string.IsNullOrEmpty(filePath))
		{
			ModelState.AddModelError("", "Faild to Save File.");
			return NotFound(itemViewFilter.File);
		}
		if (!itemViewFilter.File.FileName.EndsWith(".xlsx"))
		{
			ModelState.AddModelError("File", "Only .xlsx files are allowed.");
			return RedirectToAction("Index");
		}

		try
		{
			await ProcessExcelFileAsync(filePath);
			ViewBag.Message = "success";
			TempData["Success"] = "Successfully Migrate Data to Database";
		}
		catch (Exception ex)
		{
			// Log exception
			_logger.LogError(ex.Message);
			ModelState.AddModelError("", "Error to process the file");
			TempData["Error"] = "Error to process the file";
		}

		return RedirectToAction("Index", "ItemView");
	}

	private async Task<string> SaveFileAsync(IFormFile file)
	{
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
		}
		catch (Exception ex)
		{
			// Log exception
			_logger.LogError(ex.Message);
			return null;
		}

		return filePath;
	}

	private async Task ProcessExcelFileAsync(string filePath)
	{
		using (var stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read))
		{
			using (var reader = ExcelReaderFactory.CreateReader(stream))
			{
				do
				{
					await ProcessExcelSheetAsync(reader);
				} while (reader.NextResult());
			}
		}
	}

	private async Task ProcessExcelSheetAsync(IExcelDataReader reader)
	{
		var dataSet = reader.AsDataSet();
		var dataTable = dataSet.Tables[0];
		foreach (DataRow row in dataTable.Rows)
		{
			// skip ed header
			if (row.Table.Rows.IndexOf(row) == 0) continue;


			var itemDTO = new ItemDTO
			{
				ItemName = row[0]?.ToString(),
				Description = row[1]?.ToString(),
				ExpireDate = row[2] == DBNull.Value ? null : Convert.ToDateTime(row[2]),
				Quantity = Convert.ToInt32(row[3]),
				QuantityLow = Convert.ToInt32(row[4]),
				Availability = 	Convert.ToBoolean(row[5]),
				VolumeProperty = row[6]?.ToString(),
				PicturePath = row[7]?.ToString(),
				TypeId = 	Convert.ToInt32(row[8]),
				CategoryName = row[10]?.ToString(),
				SubCategoryName = row[11]?.ToString(),
				CityName = row[9]?.ToString(),
			};

			await AddItemToDatabaseAsync(itemDTO);
		}
	}

	private async Task AddItemToDatabaseAsync(ItemDTO itemDTO)
	{
		// Get IDs for category, subcategory, and city
		var categoryId = await GetCategoryId(itemDTO.CategoryName);
		var subCategoryId = await GetSubCategoryId(itemDTO.SubCategoryName);
		var cityId = await GetCityId(itemDTO.CityName);

		var itemCode = await _itemViewService.GenerateItemCode(categoryId, subCategoryId);

		// Map ItemDTO to Item
		DateTime today = DateTime.Now;
		var item = _mapper.Map<Item>(itemDTO);
		item.CategoryId = categoryId;
		item.SubCategoryId = subCategoryId;
		item.CityId = cityId;
		item.ItemCode = itemCode;
		item.CreateAt = today;

		// Add item to the database
		await _item.AddAsync(item);
		await _item.SaveAsync();

		// Add the item reference to the Item Views Table
		var itemView = new ItemView
		{
			ItemId = item.ItemId
		};
		await _itemView.AddAsync(itemView);
		await _itemView.SaveAsync();
	}
	

	private async Task<Guid> GetCategoryId(string categoryName)
	{
		var categoryId = (await _category.GetAll()).FirstOrDefault(c => c.CategoryName == categoryName).CategoryId;
		if (categoryId == null){
			TempData["Error"] = "Category not found";
			throw new Exception("Category not Found");

		}
		return categoryId;
	}

	private async Task<Guid> GetSubCategoryId(String subCategoryName)
	{
		var subCategoryId = (await _subCategory.GetAll()).FirstOrDefault(s => s.SubCategoryName == subCategoryName).SubCategoryId;
		if(subCategoryId == null){
			TempData["Error"] = "Subcategory not found";
			throw new Exception("Subcategory not found");
			
		}
		return  subCategoryId;
	}
	private async Task<Guid> GetCityId(String cityName)
	{
		var cityId = (await _city.GetAll()).FirstOrDefault(s => s.CityName == cityName).CityId;
		if (cityId == null)
		{
			TempData["Error"] = "City not found";
			throw new Exception("City not found");
		}
		return cityId;
	}
}
