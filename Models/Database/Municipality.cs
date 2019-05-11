﻿using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace TaxManager.Models.Database
{
    public class Municipality
    {
        public int Id { get; set; }
        [Required]
        public string Name { get; set; }
    }
}