# Application

## Endpoints

### Notes

Schedule value enum:

```
Daily = 1,
Weekly = 2,
Monthly = 3,
Yearly = 4
```

### Get Taxes

**Definition**

`GET /api/taxes?municpality=Vilnius&fromTime=2015-06-05&toTime=2019-05-05`


**Response**

- `200 Ok` on success and the filtered taxes based on give municaplity and time interval

- `404` if municpality was not found

```json
[  
   {  
      "value":0.20,
      "type":4, 
      "startDate":"2016-01-01T00:00:00", 
      "endDate":"2016-12-31T00:00:00", 
      "municipalityName":"Vilnius"
   },
   {  
      "value":0.40,
      "type":3,
      "startDate":"2016-05-01T00:00:00",
      "endDate":"2016-01-31T00:00:00",
      "municipalityName":"Vilnius"
   }
]
```

### Add a single tax on a municipality

`POST /api/taxes`

**Arguments**

- `"Value": decimal`
- `"Type": string`
- `"StartDate": dateTime`
- `"MunicipalityName": string`

**Response**

- `200 Ok` on success and the created obj

- `400` if body is invalid

```json
{  
      "value":0.20,
      "type":4, 
      "startDate":"2016-01-01T00:00:00", 
      "endDate":"2016-12-31T00:00:00", 
      "municipalityName":"Vilnius"
   }
```

### Import excel file

`POST /api/taxes/import`

Example file is named `Test.xlsx`

**Arguments**
- `file - IFormFile obj`

**Response**

- `200 Ok` on success and the created obj

- `400` if the document has an invalid row
