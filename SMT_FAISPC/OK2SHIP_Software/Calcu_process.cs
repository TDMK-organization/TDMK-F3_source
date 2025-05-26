using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OK2SHIP
{
    public static class Calcu_process
    {
        public static List<SortedDictionary<uint, double>> Group_data_Double(this IEnumerable<double> source, int totalBuckets)
        {
            var _result = new List<SortedDictionary<uint, double>>();
            for (int i = 0; i < totalBuckets; i++)
            {
                _result.Add(new SortedDictionary<uint, double>());
            }
            var min = source.Min();
            var max = source.Max();
            var bucketSize = (max - min) / totalBuckets;
            uint id = 0;
            foreach (var value in source)
            {
                int bucketIndex = 0;
                if (bucketSize > 0.0)
                {
                    bucketIndex = (int)((value - min) / bucketSize);
                    if (bucketIndex == totalBuckets)
                    {
                        bucketIndex--;
                    }
                }
                _result[bucketIndex].Add(id, value);
                id++;
            }
            return _result;
        }
        public static List<SortedDictionary<uint, double>> Group_data_String(this IEnumerable<string> source, int totalBuckets)
        {
            var _result = new List<SortedDictionary<uint, double>>();
            for (int i = 0; i < totalBuckets; i++)
            {
                _result.Add(new SortedDictionary<uint, double>());
            }
            IEnumerable<double> _source = ConvertToDouble(source);
            var min = _source.Min();
            var max = _source.Max();
            var bucketSize = (max - min) / totalBuckets;
            uint id = 0;
            foreach (var value in _source)
            {
                int bucketIndex = 0;
                if (bucketSize > 0.0)
                {
                    bucketIndex = (int)((value - min) / bucketSize);
                    if (bucketIndex == totalBuckets)
                    {
                        bucketIndex--;
                    }
                }
                _result[bucketIndex].Add(id, value);
                id++;
            }
            return _result;
        }
        public static IEnumerable<string> ConvertToString(IEnumerable<double> doubles)
        {
            return doubles.Select(ConvertToString);
        }
        public static string ConvertToString(double d)
        {
            return string.Format("{0:0.000}", d);
        }
        public static IEnumerable<double> ConvertToDouble(IEnumerable<string> src_string)
        {
            return src_string.Select(ConvertToDouble);
        }
        public static Double ConvertToDouble(string d)
        {
            double test;
            if (double.TryParse(d, out test))
            {
                return test;
            }
            else
            {
                return 0;
            }
        }
        public static int[] Bucketize_test(this IEnumerable<double> source, int totalBuckets, double bin_start, double bin_end)
        {
            var min = bin_start;
            var max = bin_end;
            var buckets = new int[totalBuckets];
            var bucketSize = (max - min) / totalBuckets;
            foreach (var value in source)
            {
                int bucketIndex = 0;
                if (bucketSize > 0.0)
                {
                    bucketIndex = (int)((value - min) / bucketSize);
                    if (bucketIndex == totalBuckets)
                    {
                        bucketIndex--;
                    }
                }
                buckets[bucketIndex]++;
            }
            return buckets;
        }
        public static List<SortedDictionary<uint, double>> Bucketize6(this IEnumerable<double> source, int totalBuckets, double bin_start, double bin_end, ref int[] item_freq)
        {
            var _result = new List<SortedDictionary<uint, double>>();
            for (int i = 0; i <= totalBuckets; i++)
            {
                _result.Add(new SortedDictionary<uint, double>());
            }
            var min = bin_start;
            var max = bin_end;
            var bucketSize = (max - min) / totalBuckets;
            uint id = 0;
            foreach (var value in source)
            {
                int bucketIndex = 0;
                if (bucketSize > 0.0)
                {
                    bucketIndex = (int)((value - min) / bucketSize);
                    if (bucketIndex == totalBuckets)
                    {
                        bucketIndex--;
                    }
                }
                if (bucketIndex <= totalBuckets)
                {
                    _result[bucketIndex].Add(id, value);
                    item_freq[bucketIndex]++;
                    id++;
                }
            }
            return _result;
        }
        public static List<SortedDictionary<uint, double>> Group_data_Freq(this IEnumerable<double> source, int totalBuckets, double bin_start, double bin_end, ref int[] item_freq)
        {
            var _result = new List<SortedDictionary<uint, double>>();
            for (int i = 0; i <= totalBuckets; i++)
            {
                _result.Add(new SortedDictionary<uint, double>());
            }
            var min = bin_start;
            var max = bin_end;
            var bucketSize = (max - min) / totalBuckets;
            uint id = 0;
            foreach (var value in source)
            {
                int bucketIndex = 0;
                if (bucketSize > 0.0)
                {
                    bucketIndex = (int)((value - min) / bucketSize);
                    if (bucketIndex == totalBuckets)
                    {
                        bucketIndex--;
                    }
                    
                }
                if ((bucketIndex <= totalBuckets) && (bucketIndex>=0))
                {
                    _result[bucketIndex].Add(id, value);
                    item_freq[bucketIndex]++;
                    id++;
                }
            }
            return _result;
        }
        public static double CPK(this IEnumerable<double> source, double USL, double LSL)
        {
            //double _result=0;
            double UL = USL;
            double LL = LSL;
            double CPKL = 0;
            double CPKU = 0;
            double stdev = CalculateStandardDeviation(source);// tg.StDev(data_arr);
            double mean = source.Average();
            CPKU = (UL - mean) / (3 * stdev);
            CPKL = (mean - LL) / (3 * stdev);
            double _result = new double[] { CPKL, CPKU }.Min();
            return _result;
        }
        public static double CalculateStandardDeviation(IEnumerable<double> values)
        {
            double standardDeviation = 0;

            if (values.Any())
            {
                // Compute the average.     
                double avg = values.Average();

                // Perform the Sum of (value-avg)_2_2.      
                double sum = values.Sum(d => Math.Pow(d - avg, 2));

                // Put it all together.      
                standardDeviation = Math.Sqrt((sum) / (values.Count() - 1));
            }

            return standardDeviation;
        }
        public static double normdist(double x, double mean, double standard_dev, bool cumalative)
        {
            if (cumalative == false)
            {
                double fact = standard_dev * Math.Sqrt(2.0 * Math.PI);
                double expo = (x - mean) * (x - mean) / (2.0 * standard_dev * standard_dev);
                return Math.Exp(-expo) / fact;
            }
            else
            {
                x = (x - mean) / standard_dev;
                if (x == 0)
                    return 0.5;
                double t = 1.0 / (1.0 + 0.2316419 * Math.Abs(x));
                double cdf = t * (1.0 / (Math.Sqrt(2.0 * Math.PI)))
                                * Math.Exp(-0.5 * x * x)
                                * (0.31938153 + t
                                * (-0.356563782 + t
                                * (1.781477937 + t
                                * (-1.821255978 + t * 1.330274429))));
                return x >= 0 ? 1.0 - cdf : cdf;
            }
        }
        public static List<SortedDictionary<uint, double>> Calcul_CPK2(Double[] data_arr, string USL, string LSL, int sigma, int count)
        {
            List<SortedDictionary<uint, double>> myData1 = new List<SortedDictionary<uint, double>>();
            List<SortedDictionary<uint, double>> myResult1 = new List<SortedDictionary<uint, double>>();
            double UL = Convert.ToDouble(USL);
            double LL = Convert.ToDouble(LSL);
            double CPKL = 0;
            double CPKU = 0;
            double stdev = CalculateStandardDeviation(data_arr);
            double mean = data_arr.Average();
            double max = data_arr.Max();
            double min = data_arr.Min();
            CPKU = (UL - mean) / (3 * stdev);
            CPKL = (mean - LL) / (3 * stdev);
            double CP = (UL - LL) / (6 * stdev);
            double CPK = new double[] { CPKL, CPKU }.Min();
            double margin = 2 * stdev;
            double mean_minus_7sig = mean - (sigma * stdev);
            double mean_plus_7sig = mean + (sigma * stdev);
            double bin_start;
            double bin_end;
            bin_start = new double[] { mean_minus_7sig, UL - margin }.Min();
            bin_end = new double[] { UL + margin, mean_minus_7sig }.Max();
            double bin_range = bin_end - bin_start;
            double Qty = count;
            double bin_step = bin_range / Qty;
            double[] bin_data = new double[count + 1];
            int[] freq_bin_data = new int[count];
            double[] modified_NormDist = new double[count];
            double[] NormDist_Bin = new double[count];
            double XiShu;
            for (int i = 0; i < count + 1; i++)
            {
                bin_data[i] = bin_start + i * bin_step;
            }
            myData1 = data_arr.Bucketize6(count, bin_start, bin_end, ref freq_bin_data);
            for (int i = 0; i < count; i++)
            {

                NormDist_Bin[i] = normdist(bin_data[i], mean, stdev, false);
            }
            XiShu = freq_bin_data.Max() / NormDist_Bin.Max();
            for (int i = 0; i < count; i++)
            {
                modified_NormDist[i] = NormDist_Bin[i] * XiShu;
                if ((modified_NormDist[i] >= freq_bin_data[i]) && (freq_bin_data[i] > 0))
                {
                    myResult1.Add(myData1[i]);
                }
            }
            return myResult1;
        }
        public static List<string> Items_OK_CPK_List(Double[] data_arr, string USL, string LSL, int sigma, int count)
        {
            List<SortedDictionary<uint, double>> myData1 = new List<SortedDictionary<uint, double>>();
            List<SortedDictionary<uint, double>> myResult1 = new List<SortedDictionary<uint, double>>();
            List<string> result = new List<string>();
            double UL = Convert.ToDouble(USL);
            double LL = Convert.ToDouble(LSL);
            double CPKL = 0;
            double CPKU = 0;
            double stdev = CalculateStandardDeviation(data_arr);
            double mean = data_arr.Average();
            double max = data_arr.Max();
            double min = data_arr.Min();
            CPKU = (UL - mean) / (3 * stdev);
            CPKL = (mean - LL) / (3 * stdev);
            double CP = (UL - LL) / (6 * stdev);
            double CPK = new double[] { CPKL, CPKU }.Min();
            double margin = 2 * stdev;
            double mean_minus_7sig = mean - (sigma * stdev);
            double mean_plus_7sig = mean + (sigma * stdev);
            double bin_start;
            double bin_end;
            bin_start = new double[] { mean_minus_7sig, UL - margin }.Min();
            bin_end = new double[] { UL + margin, mean_minus_7sig }.Max();
            double bin_range = bin_end - bin_start;
            double Qty = count;
            double bin_step = bin_range / Qty;
            double[] bin_data = new double[count + 1];
            int[] freq_bin_data = new int[count];
            double[] modified_NormDist = new double[count];
            double[] NormDist_Bin = new double[count];
            double XiShu;
            for (int i = 0; i < count + 1; i++)
            {
                bin_data[i] = bin_start + i * bin_step;
            }
            myData1 = data_arr.Bucketize6(count, bin_start, bin_end, ref freq_bin_data);
            for (int i = 0; i < count; i++)
            {

                NormDist_Bin[i] = normdist(bin_data[i], mean, stdev, false);
            }
            XiShu = freq_bin_data.Max() / NormDist_Bin.Max();
            for (int i = 0; i < count; i++)
            {
                modified_NormDist[i] = NormDist_Bin[i] * XiShu;
                if ((modified_NormDist[i] >= freq_bin_data[i]) && (freq_bin_data[i] > 0))
                {
                    myResult1.Add(myData1[i]);
                }
            }
            foreach (var item in myResult1)
            {
                if (item.Values.ToList().Count > 3)
                {
                    foreach (KeyValuePair<uint, double> x in item)
                    {
                        result.Add(x.Key.ToString());
                    }
                }
            }
            return result;
        }
        public static void calc_CPK(double[] data_arr, double UL, double LL)
        {
            double stdev = CalculateStandardDeviation(data_arr);// tg.StDev(data_arr);
            double mean = data_arr.Average();
            double max = data_arr.Max();
            double min = data_arr.Min();

            double CP = (UL - LL) / (6 * stdev);
            double CPKL = (mean - LL) / (3 * stdev);
            double CPKU = (UL - mean) / (3 * stdev);
            double CPK = new double[] { CPKL, CPKU }.Min();


        }
        
    }
}
