﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelProcessor.Models
{
    public class ExcelColumnEnum
    {
        public string TranslateIndex(int i)
        {
            switch (i)
            {
                case 1: return "A"; 
                case 2: return "B";
                case 3: return "C";
                case 4: return "D";
                case 5: return "E";
                case 6: return "F";
                case 7: return "G";
                case 8: return "H";
                case 9: return "I";
                case 10: return "J";
                case 11: return "K";
                case 12: return "L";
                case 13: return "M";
                case 14: return "N";
                case 15: return "O";
                case 16: return "P";
                case 17: return "Q";
                case 18: return "R";
                case 19: return "S";
                case 20: return "T";
                case 21: return "U";
                case 22: return "V";
                case 23: return "W";
                case 24: return "X";
                case 25: return "Y";
                case 26: return "Z";
                case 27: return "AA";
                case 28: return "AB";
                case 29: return "AC";
                case 30: return "AD";
                case 31: return "AE";
                case 32: return "AF";
                case 33: return "AG";
                case 34: return "AH";
                case 35: return "AI";
                case 36: return "AJ";
                case 37: return "AK";
                case 38: return "AL";
                case 39: return "AM";
                case 40: return "AN";
                case 41: return "AO";
                case 42: return "AP";
                case 43: return "AQ";
                case 44: return "AR";
                case 45: return "AS";
                case 46: return "AT";
                case 47: return "AU";
                case 48: return "AV";
                case 49: return "AW";
                case 50: return "AX";
                case 51: return "AY";
                case 52: return "AZ";
                case 53: return "BA";
                case 54: return "BB";
                case 55: return "BC";
                case 56: return "BD";
                case 57: return "BE";
                case 58: return "BF";
                case 59: return "BG";
                case 60: return "BH";
                case 61: return "BI";
                case 62: return "BJ";
                case 63: return "BK";
                case 64: return "BL";
                case 65: return "BM";
                case 66: return "BN";
                case 67: return "BO";
                case 68: return "BP";
                case 69: return "BQ";
                case 70: return "BR";
                default:
                    return "";
            }    
        }
    }
}
