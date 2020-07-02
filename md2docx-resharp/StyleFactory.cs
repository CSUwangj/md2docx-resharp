using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml.Wordprocessing;

namespace md2docx_resharp
{
    public class StyleFactory
    {
        public Style GenerateStyle(JsonDocument jsonDocument)
        {
            Style style = new Style { };
            return style;
        }
    }
}
