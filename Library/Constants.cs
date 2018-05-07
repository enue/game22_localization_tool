using System;
using System.Collections.Generic;
using System.Linq;

namespace Library
{
    public class Constants
    {
        readonly public static string[] Filenames = new[]
        {
            "Localization",
            "LocalizationAbility",
            "LocalizationArticle",
            "LocalizationBattle",
            "LocalizationJob",
            "LocalizationCharacter",
            "LocalizationUnit",
            "LocalizationElement",
            "LocalizationScript",
            "LocalizationCredit",
        };

        public static Dictionary<string, string> JsonExcelPaths
        {
            get
            {
                return Filenames.ToDictionary(
                    _ => "mst_develop_localize_Data/StreamingAssets/" + _ + ".json",
                    _ => _ + ".xlsx");
            }
        }

    }
}
