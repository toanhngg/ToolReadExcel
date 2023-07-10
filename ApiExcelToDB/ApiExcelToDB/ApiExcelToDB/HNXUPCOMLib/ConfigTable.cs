using Dapper;
using Microsoft.Extensions.Primitives;
using Oracle.ManagedDataAccess.Client;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Threading;
using ApiExcelToDB.Model;
using System;

namespace ApiExcelToDB.HNXUPCOMLib
{
    public class ConfigTable : ConfigApp
    {
        private readonly ConfigApp _configs;

        public ConfigTable(ConfigApp configs)
        {
            _configs = configs;
        }
        public DataSet DatTen(DataSet dataSet)
        {
            //B11,C11,D11,E11,F11,G11,H11,I11,J11
            dataSet.Tables[0].Columns[1].ColumnName = "B11";
            dataSet.Tables[0].Columns[2].ColumnName = "C11";
            dataSet.Tables[0].Columns[3].ColumnName = "D11";
            dataSet.Tables[0].Columns[4].ColumnName = "E11";
            dataSet.Tables[0].Columns[5].ColumnName = "F11";
            dataSet.Tables[0].Columns[6].ColumnName = "G11";
            dataSet.Tables[0].Columns[7].ColumnName = "H11";
            dataSet.Tables[0].Columns[8].ColumnName = "I11";
            dataSet.Tables[0].Columns[9].ColumnName = "J11";
            return dataSet;
        }
        public DataSet DatTenNY25(DataSet dataSet)
        {
            //B11,C11,D11,E11,F11,G11,H11,I11,J11
            dataSet.Tables[0].Columns[0].ColumnName = "A4";
            dataSet.Tables[0].Columns[1].ColumnName = "B4";
            dataSet.Tables[0].Columns[2].ColumnName = "C4";
            dataSet.Tables[0].Columns[3].ColumnName = "D4";
            dataSet.Tables[0].Columns[4].ColumnName = "E4";
            dataSet.Tables[0].Columns[5].ColumnName = "F4";
            dataSet.Tables[0].Columns[6].ColumnName = "G4";
            dataSet.Tables[0].Columns[7].ColumnName = "H4";
            dataSet.Tables[0].Columns[8].ColumnName = "I4";
            dataSet.Tables[0].Columns[9].ColumnName = "J4";
            dataSet.Tables[0].Columns[10].ColumnName = "K4";
            dataSet.Tables[0].Columns[11].ColumnName = "L4";
            dataSet.Tables[0].Columns[12].ColumnName = "M4";
            return dataSet;
        }
        public DataSet DatTenHNX(DataSet dataSet)
        {
            //B11,C11,D11,E11,F11,G11,H11,I11,J11
            dataSet.Tables[0].Columns[1].ColumnName = "B12";
            dataSet.Tables[0].Columns[2].ColumnName = "C12";
            dataSet.Tables[0].Columns[3].ColumnName = "D12";
            dataSet.Tables[0].Columns[4].ColumnName = "E12";
            dataSet.Tables[0].Columns[5].ColumnName = "F12";
            dataSet.Tables[0].Columns[6].ColumnName = "G12";
            dataSet.Tables[0].Columns[7].ColumnName = "H12";
            dataSet.Tables[0].Columns[8].ColumnName = "I12";
            dataSet.Tables[0].Columns[9].ColumnName = "J12";
            dataSet.Tables[0].Columns[10].ColumnName = "K12";
            dataSet.Tables[0].Columns[11].ColumnName = "L12";
            return dataSet;
        }
        //B10,C10,D10,E10,F10,G10,H10,I10,J10,
        /// <summary>
        /// K10,L10,M10,N10,O10,P10,Q10,R10
        /// </summary>
        /// <param name="dataSet"></param>
        /// <returns></returns>
        public DataSet DatTenEDO6(DataSet dataSet)
        {
            //B11,C11,D11,E11,F11,G11,H11,I11,J11
            dataSet.Tables[0].Columns[1].ColumnName = "B10";
            dataSet.Tables[0].Columns[2].ColumnName = "C10";
            dataSet.Tables[0].Columns[3].ColumnName = "D10";
            dataSet.Tables[0].Columns[4].ColumnName = "E10";
            dataSet.Tables[0].Columns[5].ColumnName = "F10";
            dataSet.Tables[0].Columns[6].ColumnName = "G10";
            dataSet.Tables[0].Columns[7].ColumnName = "H10";
            dataSet.Tables[0].Columns[8].ColumnName = "I10";
            dataSet.Tables[0].Columns[9].ColumnName = "J10";
            dataSet.Tables[0].Columns[10].ColumnName = "K10";
            dataSet.Tables[0].Columns[11].ColumnName = "L10";
            dataSet.Tables[0].Columns[12].ColumnName = "M10";
            dataSet.Tables[0].Columns[13].ColumnName = "N10";
            dataSet.Tables[0].Columns[14].ColumnName = "O10";
            dataSet.Tables[0].Columns[15].ColumnName = "P10";
            dataSet.Tables[0].Columns[16].ColumnName = "Q10";
            dataSet.Tables[0].Columns[17].ColumnName = "R10";
            return dataSet;
        }

        public DataSet DatTenUPCoM_GDNDTNN_Phien(DataSet dataSet)
        {
            //A5,B5,C5,D5,E5,F5,G5,H5,I5,J5,K5,L5,M5,N5,O5,P5,Q5
            dataSet.Tables[0].Columns[0].ColumnName = "A5";
            dataSet.Tables[0].Columns[1].ColumnName = "B5";
            dataSet.Tables[0].Columns[2].ColumnName = "C5";
            dataSet.Tables[0].Columns[3].ColumnName = "D5";
            dataSet.Tables[0].Columns[4].ColumnName = "E5";
            dataSet.Tables[0].Columns[5].ColumnName = "F5";
            dataSet.Tables[0].Columns[6].ColumnName = "G5";
            dataSet.Tables[0].Columns[7].ColumnName = "H5";
            dataSet.Tables[0].Columns[8].ColumnName = "I5";
            dataSet.Tables[0].Columns[9].ColumnName = "J5";
            dataSet.Tables[0].Columns[10].ColumnName = "K5";
            dataSet.Tables[0].Columns[11].ColumnName = "L5";
            dataSet.Tables[0].Columns[12].ColumnName = "M5";
            dataSet.Tables[0].Columns[13].ColumnName = "N5";
            dataSet.Tables[0].Columns[14].ColumnName = "O5";
            dataSet.Tables[0].Columns[15].ColumnName = "P5";
            dataSet.Tables[0].Columns[16].ColumnName = "Q5";

            return dataSet;
        }
        public DataSet DatTenUPCoM_GDNDTNN_Phien_2011(DataSet dataSet)
        {
            //A4,B4,C4,D4,E4,F4,G4,H4,I4,J4,K4,L4,M4,N4,O4,P4,Q4,
            //R4,S4,T4,U4,V4,W4,X4,Y4,Z4,AA4,AB4
            dataSet.Tables[0].Columns[0].ColumnName = "A4";
            dataSet.Tables[0].Columns[1].ColumnName = "B4";
            dataSet.Tables[0].Columns[2].ColumnName = "C4";
            dataSet.Tables[0].Columns[3].ColumnName = "D4";
            dataSet.Tables[0].Columns[4].ColumnName = "E4";
            dataSet.Tables[0].Columns[5].ColumnName = "F4";
            dataSet.Tables[0].Columns[6].ColumnName = "G4";
            dataSet.Tables[0].Columns[7].ColumnName = "H4";
            dataSet.Tables[0].Columns[8].ColumnName = "I4";
            dataSet.Tables[0].Columns[9].ColumnName = "J4";
            dataSet.Tables[0].Columns[10].ColumnName = "K4";
            dataSet.Tables[0].Columns[11].ColumnName = "L4";
            dataSet.Tables[0].Columns[12].ColumnName = "M4";
            dataSet.Tables[0].Columns[13].ColumnName = "N4";
            dataSet.Tables[0].Columns[14].ColumnName = "O4";
            dataSet.Tables[0].Columns[15].ColumnName = "P4";
            dataSet.Tables[0].Columns[16].ColumnName = "Q4";

            dataSet.Tables[0].Columns[17].ColumnName = "R4";
            dataSet.Tables[0].Columns[18].ColumnName = "S4";
            dataSet.Tables[0].Columns[19].ColumnName = "T4";
            dataSet.Tables[0].Columns[20].ColumnName = "U4";
            dataSet.Tables[0].Columns[21].ColumnName = "V4";
            dataSet.Tables[0].Columns[22].ColumnName = "W4";
            dataSet.Tables[0].Columns[23].ColumnName = "X4";
            dataSet.Tables[0].Columns[24].ColumnName = "Y4";
            dataSet.Tables[0].Columns[25].ColumnName = "Z4";
            dataSet.Tables[0].Columns[26].ColumnName = "AA4";
            dataSet.Tables[0].Columns[27].ColumnName = "AB4";


            return dataSet;
        }

        public DataSet DatTenUPCoM_KQGD_Phien(DataSet dataSet)
        {
            //A6,B6,C6,D6,E6,F6,G6,H6,I6,J6,K6,L6,M6,N6,O6,P6,Q6,R6,S6,T6,U6
            dataSet.Tables[0].Columns[0].ColumnName = "A6";
            dataSet.Tables[0].Columns[1].ColumnName = "B6";
            dataSet.Tables[0].Columns[2].ColumnName = "C6";
            dataSet.Tables[0].Columns[3].ColumnName = "D6";
            dataSet.Tables[0].Columns[4].ColumnName = "E6";
            dataSet.Tables[0].Columns[5].ColumnName = "F6";
            dataSet.Tables[0].Columns[6].ColumnName = "G6";
            dataSet.Tables[0].Columns[7].ColumnName = "H6";
            dataSet.Tables[0].Columns[8].ColumnName = "I6";
            dataSet.Tables[0].Columns[9].ColumnName = "J6";
            dataSet.Tables[0].Columns[10].ColumnName = "K6";
            dataSet.Tables[0].Columns[11].ColumnName = "L6";
            dataSet.Tables[0].Columns[12].ColumnName = "M6";
            dataSet.Tables[0].Columns[13].ColumnName = "N6";
            dataSet.Tables[0].Columns[14].ColumnName = "O6";
            dataSet.Tables[0].Columns[15].ColumnName = "P6";
            dataSet.Tables[0].Columns[16].ColumnName = "Q6";

            dataSet.Tables[0].Columns[17].ColumnName = "R6";
            dataSet.Tables[0].Columns[18].ColumnName = "S6";
            dataSet.Tables[0].Columns[19].ColumnName = "T6";
            dataSet.Tables[0].Columns[20].ColumnName = "U6";

            return dataSet;
        }

        public DataSet DatTenUPCoM_CPDKGD_Phien(DataSet dataSet)
        {
            //A4,B4,C4,D4,E4,F4,G4,H4,I4,J4,K4,L4
            dataSet.Tables[0].Columns[0].ColumnName = "A4";
            dataSet.Tables[0].Columns[1].ColumnName = "B4";
            dataSet.Tables[0].Columns[2].ColumnName = "C4";
            dataSet.Tables[0].Columns[3].ColumnName = "D4";
            dataSet.Tables[0].Columns[4].ColumnName = "E4";
            dataSet.Tables[0].Columns[5].ColumnName = "F4";
            dataSet.Tables[0].Columns[6].ColumnName = "G4";
            dataSet.Tables[0].Columns[7].ColumnName = "H4";
            dataSet.Tables[0].Columns[8].ColumnName = "I4";
            dataSet.Tables[0].Columns[9].ColumnName = "J4";
            dataSet.Tables[0].Columns[10].ColumnName = "K4";
            dataSet.Tables[0].Columns[11].ColumnName = "L4";


            return dataSet;
        }

        public DataSet DatTenUPCoM_TKCC(DataSet dataSet)
        {
            //A4,B4,C4,D4,E4,F4,G4,H4,I4,J4,K4,L4
            dataSet.Tables[0].Columns[0].ColumnName = "A4";
            dataSet.Tables[0].Columns[1].ColumnName = "B4";
            dataSet.Tables[0].Columns[2].ColumnName = "C4";
            dataSet.Tables[0].Columns[3].ColumnName = "D4";
            dataSet.Tables[0].Columns[4].ColumnName = "E4";
            dataSet.Tables[0].Columns[5].ColumnName = "F4";
            dataSet.Tables[0].Columns[6].ColumnName = "G4";


            return dataSet;
        }
        public DataSet DatTenNY_GDNDTNN_Phien(DataSet dataSet)
        {
            //B11,C11,D11,E11,F11,G11,H11,I11,J11
            dataSet.Tables[0].Columns[0].ColumnName = "A5";
            dataSet.Tables[0].Columns[1].ColumnName = "B5";
            dataSet.Tables[0].Columns[2].ColumnName = "C5";
            dataSet.Tables[0].Columns[3].ColumnName = "D5";
            dataSet.Tables[0].Columns[4].ColumnName = "E5";
            dataSet.Tables[0].Columns[5].ColumnName = "F5";
            dataSet.Tables[0].Columns[6].ColumnName = "G5";
            dataSet.Tables[0].Columns[7].ColumnName = "H5";
            dataSet.Tables[0].Columns[8].ColumnName = "I5";
            dataSet.Tables[0].Columns[9].ColumnName = "J5";
            dataSet.Tables[0].Columns[10].ColumnName = "K5";
            dataSet.Tables[0].Columns[11].ColumnName = "L5";
            dataSet.Tables[0].Columns[12].ColumnName = "M5";
            dataSet.Tables[0].Columns[13].ColumnName = "N5";
            dataSet.Tables[0].Columns[14].ColumnName = "O5";
            dataSet.Tables[0].Columns[15].ColumnName = "P5";
            dataSet.Tables[0].Columns[16].ColumnName = "Q5";
            //  dataSet.Tables[0].Columns[17].ColumnName = "R10";
            return dataSet;
        }
        //A2,B2,C2,D2,E2,F2,G2,H2,I2,J2,K2,L2,M2,N2,O2,P2
        public DataTable DatTenTT_DKGD_2011(DataTable dataTable)
        {
            //B11,C11,D11,E11,F11,G11,H11,I11,J11
            dataTable.Columns["Column0"].ColumnName = "A2";
            dataTable.Columns["Column1"].ColumnName = "B2";
            dataTable.Columns["Column2"].ColumnName = "C2";
            dataTable.Columns["Column3"].ColumnName = "D2";
            dataTable.Columns["Column4"].ColumnName = "E2";
            dataTable.Columns["Column5"].ColumnName = "F2";
            dataTable.Columns["Column6"].ColumnName = "G2";
            dataTable.Columns["Column7"].ColumnName = "H2";
            dataTable.Columns["Column8"].ColumnName = "I2";
            dataTable.Columns["Column9"].ColumnName = "J2";
            dataTable.Columns["Column10"].ColumnName = "K2";
            dataTable.Columns["Column11"].ColumnName = "L2";
            dataTable.Columns["Column12"].ColumnName = "M2";
            dataTable.Columns["Column13"].ColumnName = "N2";
            dataTable.Columns["Column14"].ColumnName = "O2";
            dataTable.Columns["Column15"].ColumnName = "P2";

            return dataTable;
        }
        public DataTable DatTenTT_DKGD_2011SS(DataTable dataTable)
        {
            //B11,C11,D11,E11,F11,G11,H11,I11,J11
            dataTable.Columns["Column0"].ColumnName = "A2";
            dataTable.Columns["Column1"].ColumnName = "B2";
            dataTable.Columns["Column2"].ColumnName = "C2";
            dataTable.Columns["Column3"].ColumnName = "D2";
            dataTable.Columns["Column4"].ColumnName = "E2";
            dataTable.Columns["Column5"].ColumnName = "F2";
            dataTable.Columns["Column6"].ColumnName = "G2";
            dataTable.Columns["Column7"].ColumnName = "H2";
            dataTable.Columns["Column8"].ColumnName = "I2";
            dataTable.Columns["Column9"].ColumnName = "J2";
            dataTable.Columns["Column10"].ColumnName = "K2";
            dataTable.Columns["Column11"].ColumnName = "L2";
            dataTable.Columns["Column12"].ColumnName = "M2";
            dataTable.Columns["Column13"].ColumnName = "N2";
            dataTable.Columns["Column14"].ColumnName = "O2";
            //   dataTable.Columns["Column15"].ColumnName = "P2";

            return dataTable;
        }
        public DataTable DatTenTT_DKGD_2011_21(DataTable dataTable)
        {
            //B11,C11,D11,E11,F11,G11,H11,I11,J11
            dataTable.Columns["Column0"].ColumnName = "A2";
            dataTable.Columns["Column1"].ColumnName = "B2";
            dataTable.Columns["Column2"].ColumnName = "C2";
            dataTable.Columns["Column3"].ColumnName = "D2";
            dataTable.Columns["Column4"].ColumnName = "E2";
            dataTable.Columns["Column5"].ColumnName = "F2";
            dataTable.Columns["Column6"].ColumnName = "G2";
            dataTable.Columns["Column7"].ColumnName = "H2";
            dataTable.Columns["Column8"].ColumnName = "I2";
            dataTable.Columns["Column9"].ColumnName = "J2";
            dataTable.Columns["Column10"].ColumnName = "K2";
            dataTable.Columns["Column11"].ColumnName = "L2";
            dataTable.Columns["Column12"].ColumnName = "M2";


            return dataTable;
        }

        public DataTable DatTenNDTNN_2011(DataTable dataTable)
        {
            //A3,B3,C3,D3,E3,F3,G3,H3,I3,J3,K3,L3,M3,N3,O3,P3
            dataTable.Columns["Column0"].ColumnName = "A3";
            dataTable.Columns["Column1"].ColumnName = "B3";
            dataTable.Columns["Column2"].ColumnName = "C3";
            dataTable.Columns["Column3"].ColumnName = "D3";
            dataTable.Columns["Column4"].ColumnName = "E3";
            dataTable.Columns["Column5"].ColumnName = "F3";
            dataTable.Columns["Column6"].ColumnName = "G3";
            dataTable.Columns["Column7"].ColumnName = "H3";
            dataTable.Columns["Column8"].ColumnName = "I3";
            dataTable.Columns["Column9"].ColumnName = "J3";
            dataTable.Columns["Column10"].ColumnName = "K3";
            dataTable.Columns["Column11"].ColumnName = "L3";
            dataTable.Columns["Column12"].ColumnName = "M3";
            dataTable.Columns["Column13"].ColumnName = "N3";
            dataTable.Columns["Column14"].ColumnName = "O3";
            dataTable.Columns["Column15"].ColumnName = "P3";

            return dataTable;
        }

        public DataTable DatTenKQGD_2011(DataTable dataTable)
        {
            //A3,B3,C3,D3,E3,F3,G3,H3,I3,J3,K3,L3,M3,N3,O3,P3
            dataTable.Columns["Column0"].ColumnName = "A2";
            dataTable.Columns["Column1"].ColumnName = "B2";
            dataTable.Columns["Column2"].ColumnName = "C2";
            dataTable.Columns["Column3"].ColumnName = "D2";
            dataTable.Columns["Column4"].ColumnName = "E2";
            dataTable.Columns["Column5"].ColumnName = "F2";
            dataTable.Columns["Column6"].ColumnName = "G2";
            dataTable.Columns["Column7"].ColumnName = "H2";
            dataTable.Columns["Column8"].ColumnName = "I2";
            dataTable.Columns["Column9"].ColumnName = "J2";
            dataTable.Columns["Column10"].ColumnName = "K2";
            dataTable.Columns["Column11"].ColumnName = "L2";
            dataTable.Columns["Column12"].ColumnName = "M2";
            dataTable.Columns["Column13"].ColumnName = "N2";
            dataTable.Columns["Column14"].ColumnName = "O2";
            dataTable.Columns["Column15"].ColumnName = "P2";
            dataTable.Columns["Column16"].ColumnName = "Q2";

            return dataTable;
        }
        public DataTable DatTenTop10CK_KLGDL(DataTable dataTable)
        {
            //A3,B3,C3,D3,E3,F3,G3,H3,I3,J3,K3,L3,M3,N3,O3,P3

            dataTable.Columns["Column6"].ColumnName = "G3";
            dataTable.Columns["Column7"].ColumnName = "H3";
            dataTable.Columns["Column8"].ColumnName = "I3";
            dataTable.Columns["Column9"].ColumnName = "J3";
            dataTable.Columns["Column10"].ColumnName = "K3";

            return dataTable;
        }
        public DataTable DatTenTop10CK_GTGDL(DataTable dataTable)
        {
            //A3,B3,C3,D3,E3,F3,G3,H3,I3,J3,K3,L3,M3,N3,O3,P3

            dataTable.Columns["Column2"].ColumnName = "C3";
            dataTable.Columns["Column3"].ColumnName = "D3";
            dataTable.Columns["Column4"].ColumnName = "E3";


            return dataTable;
        }
        public DataTable DatTenTop10CK_GTGDL_2(DataTable dataTable)
        {
            //A3,B3,C3,D3,E3,F3,G3,H3,I3,J3,K3,L3,M3,N3,O3,P3

            dataTable.Columns["Column1"].ColumnName = "C3";
            dataTable.Columns["Column2"].ColumnName = "D3";
            dataTable.Columns["Column3"].ColumnName = "E3";


            return dataTable;
        }
        public DataTable DatTenTop10CK_GTGDL_2010(DataTable dataTable)
        {
            //A3,B3,C3,D3,E3,F3,G3,H3,I3,J3,K3,L3,M3,N3,O3,P3

            dataTable.Columns["Column0"].ColumnName = "A3";
            dataTable.Columns["Column1"].ColumnName = "B3";
            dataTable.Columns["Column2"].ColumnName = "C3";


            return dataTable;
        }
        public DataTable DatTenTop10CK_KLGDL_2010(DataTable dataTable)
        {
            //A3,B3,C3,D3,E3,F3,G3,H3,I3,J3,K3,L3,M3,N3,O3,P3

            dataTable.Columns["Column4"].ColumnName = "E3";
            dataTable.Columns["Column5"].ColumnName = "F3";
            dataTable.Columns["Column6"].ColumnName = "G3";

            return dataTable;
        }
        public DataTable DatTenTop10CP_GTNYL(DataTable dataTable)
        {


            dataTable.Columns["Column14"].ColumnName = "O3";
            dataTable.Columns["Column15"].ColumnName = "P3";
            dataTable.Columns["Column16"].ColumnName = "Q3";
            dataTable.Columns["Column17"].ColumnName = "R3";


            return dataTable;
        }
        public DataTable DatTenTop10CP_GTNYL_2(DataTable dataTable)
        {


            dataTable.Columns["Column13"].ColumnName = "O3";
            dataTable.Columns["Column14"].ColumnName = "P3";
            dataTable.Columns["Column15"].ColumnName = "Q3";
            dataTable.Columns["Column16"].ColumnName = "R3";


            return dataTable;
        }
        public DataTable DatTenTop10CP_GTNYL_2010_1(DataTable dataTable)
        {


            dataTable.Columns["Column8"].ColumnName = "I3";
            dataTable.Columns["Column9"].ColumnName = "J3";
            dataTable.Columns["Column10"].ColumnName = "K3";
            dataTable.Columns["Column11"].ColumnName = "L3";


            return dataTable;
        }
        public DataTable DatTenTop10CK_TANGGIA(DataTable dataTable)
        {


            dataTable.Columns["Column23"].ColumnName = "X3";
            dataTable.Columns["Column24"].ColumnName = "Y3";
            dataTable.Columns["Column25"].ColumnName = "Z3";
            dataTable.Columns["Column26"].ColumnName = "AA3";


            return dataTable;
        }
        public DataTable DatTenTop10CK_TANGGIA_2(DataTable dataTable)
        {


            dataTable.Columns["Column21"].ColumnName = "X3";
            dataTable.Columns["Column22"].ColumnName = "Y3";
            dataTable.Columns["Column23"].ColumnName = "Z3";
            dataTable.Columns["Column24"].ColumnName = "AA3";


            return dataTable;
        }
        public DataTable DatTenTop10CK_TANGGIA_2010(DataTable dataTable)
        {


            dataTable.Columns["Column13"].ColumnName = "N3";
            dataTable.Columns["Column14"].ColumnName = "O3";
            dataTable.Columns["Column15"].ColumnName = "P3";
            dataTable.Columns["Column16"].ColumnName = "Q3";


            return dataTable;
        }
        public DataTable DatTenTop10CK_GIAMGIA(DataTable dataTable)
        {


            dataTable.Columns["Column28"].ColumnName = "AC3";
            dataTable.Columns["Column29"].ColumnName = "AD3";
            dataTable.Columns["Column30"].ColumnName = "AE3";
            dataTable.Columns["Column31"].ColumnName = "AF3";


            return dataTable;
        }
        public DataTable DatTenTop10CK_GIAMGIA_2010(DataTable dataTable)
        {


            dataTable.Columns["Column18"].ColumnName = "S3";
            dataTable.Columns["Column19"].ColumnName = "T3";
            dataTable.Columns["Column20"].ColumnName = "U3";
            dataTable.Columns["Column21"].ColumnName = "V3";


            return dataTable;
        }
        public DataTable DatTenTop10CK_GIAMGIA_2(DataTable dataTable)
        {


            dataTable.Columns["Column27"].ColumnName = "AC3";
            dataTable.Columns["Column28"].ColumnName = "AD3";
            dataTable.Columns["Column29"].ColumnName = "AE3";
            dataTable.Columns["Column30"].ColumnName = "AF3";


            return dataTable;
        }
        public DataTable DatTenChi_Tieu_2011(DataTable dataTable)
        {


            dataTable.Columns["Column35"].ColumnName = "AJ3";
            dataTable.Columns["Column36"].ColumnName = "AK3";
            dataTable.Columns["Column37"].ColumnName = "AL3";



            return dataTable;
        }

        public DataTable DatTenChi_Tieu_2011_2(DataTable dataTable)
        {


            dataTable.Columns["Column34"].ColumnName = "AJ3";
            dataTable.Columns["Column35"].ColumnName = "AK3";
            dataTable.Columns["Column36"].ColumnName = "AL3";



            return dataTable;
        }
        public DataTable DatTenTop10CK_NDTNN(DataTable dataTable)
        {


            dataTable.Columns["Column39"].ColumnName = "AN3";
            dataTable.Columns["Column40"].ColumnName = "AO3";
            dataTable.Columns["Column41"].ColumnName = "AP3";
            dataTable.Columns["Column42"].ColumnName = "AQ3";


            return dataTable;
        }
        public DataTable DatTenTop10CK_NDTNN_2(DataTable dataTable)
        {


            dataTable.Columns["Column38"].ColumnName = "AN3";
            dataTable.Columns["Column39"].ColumnName = "AO3";
            dataTable.Columns["Column40"].ColumnName = "AP3";
            dataTable.Columns["Column41"].ColumnName = "AQ3";


            return dataTable;
        }

        public DataTable DatTenKLGD_TOP2011_MR(DataTable dataTable)
        {

            //B16,C16,D16,E16,F16,G16
            dataTable.Columns["Column1"].ColumnName = "B16";
            dataTable.Columns["C3"].ColumnName = "C16";
            dataTable.Columns["D3"].ColumnName = "D16";
            dataTable.Columns["E3"].ColumnName = "E16";
            dataTable.Columns["Column5"].ColumnName = "F16";
            dataTable.Columns["G3"].ColumnName = "G16";


            return dataTable;
        }
        public DataTable DatTenKLGD_TOP2011_MR_2(DataTable dataTable)
        {

            //B16,C16,D16,E16,F16,G16
            dataTable.Columns["C3"].ColumnName = "B16";
            dataTable.Columns["D3"].ColumnName = "C16";
            dataTable.Columns["E3"].ColumnName = "D16";
            dataTable.Columns["Column4"].ColumnName = "E16";
            dataTable.Columns["Column5"].ColumnName = "F16";
            dataTable.Columns["G3"].ColumnName = "G16";


            return dataTable;
        }
        public DataTable DatTenGTGD_TOP2011_MR(DataTable dataTable)
        {

            //K16,L16,M16,N16,O16,P16
            dataTable.Columns["K3"].ColumnName = "K16";
            dataTable.Columns["Column11"].ColumnName = "L16";
            dataTable.Columns["Column12"].ColumnName = "M16";
            dataTable.Columns["Column13"].ColumnName = "N16";
            dataTable.Columns["O3"].ColumnName = "O16";
            dataTable.Columns["P3"].ColumnName = "P16";


            return dataTable;
        }
        public DataTable DatTenGTGD_TOP2011_MR_2(DataTable dataTable)
        {

            //K16,L16,M16,N16,O16,P16
            dataTable.Columns["K3"].ColumnName = "K16";
            dataTable.Columns["Column11"].ColumnName = "L16";
            dataTable.Columns["Column12"].ColumnName = "M16";
            dataTable.Columns["O3"].ColumnName = "N16";
            dataTable.Columns["P3"].ColumnName = "O16";
            dataTable.Columns["Q3"].ColumnName = "P16";


            return dataTable;
        }
        public DataTable DatTenTangGiam_TOP2011_MR(DataTable dataTable)
        {

            //U16,V16,W16,X16,Y16,Z16,AA16,AB16,AC16
            dataTable.Columns["Column20"].ColumnName = "U16";
            dataTable.Columns["Column21"].ColumnName = "V16";
            dataTable.Columns["Column22"].ColumnName = "W16";
            dataTable.Columns["X3"].ColumnName = "X16";
            dataTable.Columns["Y3"].ColumnName = "Y16";
            dataTable.Columns["Z3"].ColumnName = "Z16";
            dataTable.Columns["AA3"].ColumnName = "AA16";
            dataTable.Columns["Column27"].ColumnName = "AB16";
            dataTable.Columns["AC3"].ColumnName = "AC16";


            return dataTable;
        }
        public DataTable DatTenTangGiam_TOP2011_MR_SS(DataTable dataTable)
        {

            //U16,V16,W16,X16,Y16,Z16,AA16,AB16,AC16
            dataTable.Columns["Column19"].ColumnName = "U16";
            dataTable.Columns["Column20"].ColumnName = "V16";
            dataTable.Columns["Column21"].ColumnName = "W16";
            dataTable.Columns["Column22"].ColumnName = "X16";
            dataTable.Columns["X3"].ColumnName = "Y16";
            dataTable.Columns["Y3"].ColumnName = "Z16";
            dataTable.Columns["Z3"].ColumnName = "AA16";
            dataTable.Columns["AA3"].ColumnName = "AB16";
            dataTable.Columns["AC3"].ColumnName = "AC16";


            return dataTable;
        }
        public DataTable DatTenTangGiam_TOP2011_MR_2(DataTable dataTable)
        {

            //U16,V16,W16,X16,Y16,Z16,AA16,AB16,AC16
            dataTable.Columns["Column19"].ColumnName = "U16";
            dataTable.Columns["Column20"].ColumnName = "V16";
            dataTable.Columns["X3"].ColumnName = "W16";
            dataTable.Columns["Y3"].ColumnName = "X16";
            dataTable.Columns["Z3"].ColumnName = "Y16";
            dataTable.Columns["AA3"].ColumnName = "Z16";
            dataTable.Columns["Column25"].ColumnName = "AA16";
            dataTable.Columns["Column26"].ColumnName = "AB16";
            dataTable.Columns["AC3"].ColumnName = "AC16";


            return dataTable;
        }

        public DataTable DatTenCKNTDNN_TOP2011_MR(DataTable dataTable)
        {

            //AN16,AO16,AP16,AQ16
            dataTable.Columns["AN3"].ColumnName = "AN16";
            dataTable.Columns["AO3"].ColumnName = "AO16";
            dataTable.Columns["AP3"].ColumnName = "AP16";
            dataTable.Columns["AQ3"].ColumnName = "AQ16";
            dataTable.Columns["Column43"].ColumnName = "AR16";
            dataTable.Columns["Column44"].ColumnName = "AS16";
            return dataTable;
        }
        public DataTable DatTenCKNTDNN_TOP2011_MR_2(DataTable dataTable)
        {

            //AN16,AO16,AP16,AQ16
            dataTable.Columns["AN3"].ColumnName = "AN16";
            dataTable.Columns["AO3"].ColumnName = "AO16";
            dataTable.Columns["AP3"].ColumnName = "AP16";
            dataTable.Columns["AQ3"].ColumnName = "AQ16";
            // dataTable.Columns["Column43"].ColumnName = "AR16";
            // dataTable.Columns["Column44"].ColumnName = "AS16";
            return dataTable;
        }



        /* public DataTable DatTenTop10CK_KLGDL(DataTable dataTable)
         {
             //A3,B3,C3,D3,E3,F3,G3,H3,I3,J3,K3,L3,M3,N3,O3,P3

             dataTable.Columns["Column2"].ColumnName = "C3";
             dataTable.Columns["Column3"].ColumnName = "D3";
             dataTable.Columns["Column4"].ColumnName = "E3";

             return dataTable;
         }*/
        public DataTable DatTenKQGDTH_2011(DataTable dataTable)
        {
            //A3,B3,C3,D3,E3,F3,G3,H3,I3,J3,K3,L3,M3,N3,O3,P3
            dataTable.Columns["Column0"].ColumnName = "A6";
            dataTable.Columns["Column1"].ColumnName = "B6";
            dataTable.Columns["Column2"].ColumnName = "C6";
            dataTable.Columns["Column3"].ColumnName = "D6";
            dataTable.Columns["Column4"].ColumnName = "E6";
            dataTable.Columns["Column5"].ColumnName = "F6";
            dataTable.Columns["Column6"].ColumnName = "G6";
            dataTable.Columns["Column7"].ColumnName = "H6";
            dataTable.Columns["Column8"].ColumnName = "I6";
            dataTable.Columns["Column9"].ColumnName = "J6";
            dataTable.Columns["Column10"].ColumnName = "K6";
            dataTable.Columns["Column11"].ColumnName = "L6";

            return dataTable;
        }

        public DataTable DatTenTH_DATLENH_2011(DataTable dataTable)
        {
            //A6,B6,C6,D6,E6,F6,G6,H6,I6,J6,K6,L6,M6,N6,O6
            dataTable.Columns["Column0"].ColumnName = "A6";
            dataTable.Columns["Column1"].ColumnName = "B6";
            dataTable.Columns["Column2"].ColumnName = "C6";
            dataTable.Columns["Column3"].ColumnName = "D6";
            dataTable.Columns["Column4"].ColumnName = "E6";
            dataTable.Columns["Column5"].ColumnName = "F6";
            dataTable.Columns["Column6"].ColumnName = "G6";
            dataTable.Columns["Column7"].ColumnName = "H6";
            dataTable.Columns["Column8"].ColumnName = "I6";
            dataTable.Columns["Column9"].ColumnName = "J6";
            dataTable.Columns["Column10"].ColumnName = "K6";
            dataTable.Columns["Column11"].ColumnName = "L6";
            dataTable.Columns["Column12"].ColumnName = "M6";
            dataTable.Columns["Column13"].ColumnName = "N6";
            dataTable.Columns["Column14"].ColumnName = "O6";


            return dataTable;
        }
        public DataTable DatTenTH_GDTRAIPHIEU_2010(DataTable dataTable)
        {
            //A6,B6,C6,D6,E6,F6,G6,H6,I6,J6,K6,L6,M6,N6,O6
            dataTable.Columns["Column0"].ColumnName = "A2";
            dataTable.Columns["Column1"].ColumnName = "B2";
            dataTable.Columns["Column2"].ColumnName = "C2";
            dataTable.Columns["Column3"].ColumnName = "D2";
            dataTable.Columns["Column4"].ColumnName = "E2";
            dataTable.Columns["Column5"].ColumnName = "F2";
            dataTable.Columns["Column6"].ColumnName = "G2";
            dataTable.Columns["Column7"].ColumnName = "H2";



            return dataTable;
        }
        public DataTable DatTenTH_GDTP_NDTNN_2010(DataTable dataTable)
        {
            //A6,B6,C6,D6,E6,F6,G6,H6,I6,J6,K6,L6,M6,N6,O6
            dataTable.Columns["Column0"].ColumnName = "A3";
            dataTable.Columns["Column1"].ColumnName = "B3";
            dataTable.Columns["Column2"].ColumnName = "C3";
            dataTable.Columns["Column3"].ColumnName = "D3";
            dataTable.Columns["Column4"].ColumnName = "E3";
            dataTable.Columns["Column5"].ColumnName = "F3";
            dataTable.Columns["Column6"].ColumnName = "G3";
            dataTable.Columns["Column7"].ColumnName = "H3";
            dataTable.Columns["Column8"].ColumnName = "I3";
            dataTable.Columns["Column9"].ColumnName = "J3";
            dataTable.Columns["Column10"].ColumnName = "K3";
            dataTable.Columns["Column11"].ColumnName = "L3";
            dataTable.Columns["Column12"].ColumnName = "M3";
            dataTable.Columns["Column13"].ColumnName = "N3";
            dataTable.Columns["Column14"].ColumnName = "O3";
            dataTable.Columns["Column15"].ColumnName = "P3";




            return dataTable;
        }
        public DataSet DatTenNY24(DataSet dataSet)
        {
            //B11,C11,D11,E11,F11,G11,H11,I11,J11
            dataSet.Tables[0].Columns[0].ColumnName = "A5";
            dataSet.Tables[0].Columns[1].ColumnName = "B5";
            dataSet.Tables[0].Columns[2].ColumnName = "C5";
            dataSet.Tables[0].Columns[3].ColumnName = "D5";
            dataSet.Tables[0].Columns[4].ColumnName = "E5";
            dataSet.Tables[0].Columns[5].ColumnName = "F5";
            dataSet.Tables[0].Columns[6].ColumnName = "G5";
            dataSet.Tables[0].Columns[7].ColumnName = "H5";
            dataSet.Tables[0].Columns[8].ColumnName = "I5";
            dataSet.Tables[0].Columns[9].ColumnName = "J5";
            dataSet.Tables[0].Columns[10].ColumnName = "K5";
            dataSet.Tables[0].Columns[11].ColumnName = "L5";
            dataSet.Tables[0].Columns[12].ColumnName = "M5";
            dataSet.Tables[0].Columns[13].ColumnName = "N5";
            dataSet.Tables[0].Columns[14].ColumnName = "O5";
            dataSet.Tables[0].Columns[15].ColumnName = "P5";
            dataSet.Tables[0].Columns[16].ColumnName = "Q5";

            return dataSet;
        }
        public DataSet DatTenEDO4(DataSet dataSet)
        {
            //B10,C10,D10,E10,F10,G10,H10,I10,J10,K10,L10,M10,N10,O10,P10,Q10,R10,S10
            dataSet.Tables[0].Columns[1].ColumnName = "B10";
            dataSet.Tables[0].Columns[2].ColumnName = "C10";
            dataSet.Tables[0].Columns[3].ColumnName = "D10";
            dataSet.Tables[0].Columns[4].ColumnName = "E10";
            dataSet.Tables[0].Columns[5].ColumnName = "F10";
            dataSet.Tables[0].Columns[6].ColumnName = "G10";
            dataSet.Tables[0].Columns[7].ColumnName = "H10";
            dataSet.Tables[0].Columns[8].ColumnName = "I10";
            dataSet.Tables[0].Columns[9].ColumnName = "J10";
            dataSet.Tables[0].Columns[10].ColumnName = "K10";
            dataSet.Tables[0].Columns[11].ColumnName = "L10";
            dataSet.Tables[0].Columns[12].ColumnName = "M10";
            dataSet.Tables[0].Columns[13].ColumnName = "N10";
            dataSet.Tables[0].Columns[14].ColumnName = "O10";
            dataSet.Tables[0].Columns[15].ColumnName = "P10";
            dataSet.Tables[0].Columns[16].ColumnName = "Q10";
            dataSet.Tables[0].Columns[17].ColumnName = "R10";
            dataSet.Tables[0].Columns[18].ColumnName = "S10";
            return dataSet;
        }
        public DataSet DatTenNY_EOD4_View(DataSet dataSet)
        {
            //B10,C10,D10,E10,F10,G10,H10,I10,J10,K10,L10,M10,N10,O10,P10,Q10,R10,S10
            dataSet.Tables[0].Columns[0].ColumnName = "A5";
            dataSet.Tables[0].Columns[1].ColumnName = "B5";
            dataSet.Tables[0].Columns[2].ColumnName = "C5";
            dataSet.Tables[0].Columns[3].ColumnName = "D5";
            dataSet.Tables[0].Columns[4].ColumnName = "E5";
            dataSet.Tables[0].Columns[5].ColumnName = "F5";
            dataSet.Tables[0].Columns[6].ColumnName = "G5";
            dataSet.Tables[0].Columns[7].ColumnName = "H5";
            dataSet.Tables[0].Columns[8].ColumnName = "I5";
            dataSet.Tables[0].Columns[9].ColumnName = "J5";
            dataSet.Tables[0].Columns[10].ColumnName = "K5";
            dataSet.Tables[0].Columns[11].ColumnName = "L5";
            dataSet.Tables[0].Columns[12].ColumnName = "M5";
            dataSet.Tables[0].Columns[13].ColumnName = "N5";
            dataSet.Tables[0].Columns[14].ColumnName = "O5";
            dataSet.Tables[0].Columns[15].ColumnName = "P5";
            dataSet.Tables[0].Columns[16].ColumnName = "Q5";
            dataSet.Tables[0].Columns[17].ColumnName = "R5";
            //  dataSet.Tables[0].Columns[18].ColumnName = "S10";
            return dataSet;
        }
        public DataSet DatTenNY23(DataSet dataSet)
        {
            //B10,C10,D10,E10,F10,G10,H10,I10,J10,K10,L10,M10,N10,O10,P10,Q10,R10,S10
            dataSet.Tables[0].Columns[0].ColumnName = "A5";
            dataSet.Tables[0].Columns[1].ColumnName = "B5";
            dataSet.Tables[0].Columns[2].ColumnName = "C5";
            dataSet.Tables[0].Columns[3].ColumnName = "D5";
            dataSet.Tables[0].Columns[4].ColumnName = "E5";
            dataSet.Tables[0].Columns[5].ColumnName = "F5";
            dataSet.Tables[0].Columns[6].ColumnName = "G5";
            dataSet.Tables[0].Columns[7].ColumnName = "H5";
            dataSet.Tables[0].Columns[8].ColumnName = "I5";
            dataSet.Tables[0].Columns[9].ColumnName = "J5";
            dataSet.Tables[0].Columns[10].ColumnName = "K5";
            dataSet.Tables[0].Columns[11].ColumnName = "L5";
            dataSet.Tables[0].Columns[12].ColumnName = "M5";
            dataSet.Tables[0].Columns[13].ColumnName = "N5";
            dataSet.Tables[0].Columns[14].ColumnName = "O5";
            dataSet.Tables[0].Columns[15].ColumnName = "P5";
            dataSet.Tables[0].Columns[16].ColumnName = "Q5";
            dataSet.Tables[0].Columns[17].ColumnName = "R5";

            return dataSet;
        }
        public DataSet DatTenEDO1(DataSet dataSet)
        {
            //B8,C8,D8,E8,F8,G8,H8,I8,J8,K8,L8,M8,N8,O8,P8,Q8,R8,S8,T8,U8,V8,W8,X8,Y8,Z8
            dataSet.Tables[0].Columns[1].ColumnName = "B8";
            dataSet.Tables[0].Columns[2].ColumnName = "C8";
            dataSet.Tables[0].Columns[3].ColumnName = "D8";
            dataSet.Tables[0].Columns[4].ColumnName = "E8";
            dataSet.Tables[0].Columns[5].ColumnName = "F8";
            dataSet.Tables[0].Columns[6].ColumnName = "G8";
            dataSet.Tables[0].Columns[7].ColumnName = "H8";
            dataSet.Tables[0].Columns[8].ColumnName = "I8";
            dataSet.Tables[0].Columns[9].ColumnName = "J8";
            dataSet.Tables[0].Columns[10].ColumnName = "K8";
            dataSet.Tables[0].Columns[11].ColumnName = "L8";
            dataSet.Tables[0].Columns[12].ColumnName = "M8";
            dataSet.Tables[0].Columns[13].ColumnName = "N8";
            dataSet.Tables[0].Columns[14].ColumnName = "O8";
            dataSet.Tables[0].Columns[15].ColumnName = "P8";
            dataSet.Tables[0].Columns[16].ColumnName = "Q8";
            dataSet.Tables[0].Columns[17].ColumnName = "R8";
            dataSet.Tables[0].Columns[18].ColumnName = "S8";
            dataSet.Tables[0].Columns[19].ColumnName = "T8";
            dataSet.Tables[0].Columns[20].ColumnName = "U8";
            dataSet.Tables[0].Columns[21].ColumnName = "V8";
            dataSet.Tables[0].Columns[22].ColumnName = "W8";
            dataSet.Tables[0].Columns[23].ColumnName = "X8";
            dataSet.Tables[0].Columns[24].ColumnName = "Y8";
            dataSet.Tables[0].Columns[25].ColumnName = "Z8";

            return dataSet;
        }
        public DataSet DatTenNY21(DataSet dataSet)
        {

            dataSet.Tables[0].Columns[0].ColumnName = "A6";
            dataSet.Tables[0].Columns[1].ColumnName = "B6";
            dataSet.Tables[0].Columns[2].ColumnName = "C6";
            dataSet.Tables[0].Columns[3].ColumnName = "D6";
            dataSet.Tables[0].Columns[4].ColumnName = "E6";
            dataSet.Tables[0].Columns[5].ColumnName = "F6";
            dataSet.Tables[0].Columns[6].ColumnName = "G6";
            dataSet.Tables[0].Columns[7].ColumnName = "H6";
            dataSet.Tables[0].Columns[8].ColumnName = "I6";
            dataSet.Tables[0].Columns[9].ColumnName = "J6";
            dataSet.Tables[0].Columns[10].ColumnName = "K6";
            dataSet.Tables[0].Columns[11].ColumnName = "L6";
            dataSet.Tables[0].Columns[12].ColumnName = "M6";
            dataSet.Tables[0].Columns[13].ColumnName = "N6";
            dataSet.Tables[0].Columns[14].ColumnName = "O6";
            dataSet.Tables[0].Columns[15].ColumnName = "P6";
            dataSet.Tables[0].Columns[16].ColumnName = "Q6";
            dataSet.Tables[0].Columns[17].ColumnName = "R6";
            dataSet.Tables[0].Columns[18].ColumnName = "S6";
            dataSet.Tables[0].Columns[19].ColumnName = "T6";
            dataSet.Tables[0].Columns[20].ColumnName = "U6";
            dataSet.Tables[0].Columns[21].ColumnName = "V6";
            dataSet.Tables[0].Columns[22].ColumnName = "W6";


            return dataSet;
        }

        public DataSet DatTenEDO1S(DataSet dataSet)
        {
            //B8,C8,D8,E8,F8,G8,H8,I8,J8,K8,L8,M8,N8,O8,P8,Q8,R8,S8,T8,U8,V8,W8,X8,Y8,Z8
            dataSet.Tables[0].Columns[1].ColumnName = "B8";
            dataSet.Tables[0].Columns[2].ColumnName = "C8";
            dataSet.Tables[0].Columns[3].ColumnName = "D8";
            dataSet.Tables[0].Columns[4].ColumnName = "E8";
            dataSet.Tables[0].Columns[5].ColumnName = "F8";
            dataSet.Tables[0].Columns[6].ColumnName = "G8";
            dataSet.Tables[0].Columns[7].ColumnName = "H8";
            dataSet.Tables[0].Columns[8].ColumnName = "I8";
            dataSet.Tables[0].Columns[9].ColumnName = "J8";
            dataSet.Tables[0].Columns[10].ColumnName = "K8";
            dataSet.Tables[0].Columns[11].ColumnName = "L8";
            dataSet.Tables[0].Columns[12].ColumnName = "M8";
            dataSet.Tables[0].Columns[13].ColumnName = "N8";
            dataSet.Tables[0].Columns[14].ColumnName = "O8";
            dataSet.Tables[0].Columns[15].ColumnName = "P8";
            dataSet.Tables[0].Columns[16].ColumnName = "Q8";
            dataSet.Tables[0].Columns[17].ColumnName = "R8";
            dataSet.Tables[0].Columns[18].ColumnName = "S8";
            dataSet.Tables[0].Columns[19].ColumnName = "T8";
            dataSet.Tables[0].Columns[20].ColumnName = "U8";
            dataSet.Tables[0].Columns[21].ColumnName = "V8";
            dataSet.Tables[0].Columns[22].ColumnName = "W8";
            dataSet.Tables[0].Columns[23].ColumnName = "X8";
            dataSet.Tables[0].Columns[24].ColumnName = "Y8";
            dataSet.Tables[0].Columns[25].ColumnName = "Z8";
            dataSet.Tables[0].Columns[26].ColumnName = "AA8";
            dataSet.Tables[0].Columns[27].ColumnName = "AB8";

            return dataSet;
        }

        public DataSet DatTenEDO1_HNX(DataSet dataSet)
        {
            //B8,C8,D8,E8,F8,G8,H8,I8,J8,K8,L8,M8,N8,O8,P8,Q8,R8,S8,T8,U8,V8,W8,X8,Y8,Z8
            dataSet.Tables[0].Columns[1].ColumnName = "B8";
            dataSet.Tables[0].Columns[2].ColumnName = "C8";
            dataSet.Tables[0].Columns[3].ColumnName = "D8";
            dataSet.Tables[0].Columns[4].ColumnName = "E8";
            dataSet.Tables[0].Columns[5].ColumnName = "F8";
            dataSet.Tables[0].Columns[6].ColumnName = "G8";
            dataSet.Tables[0].Columns[7].ColumnName = "H8";
            dataSet.Tables[0].Columns[8].ColumnName = "I8";
            dataSet.Tables[0].Columns[9].ColumnName = "J8";
            dataSet.Tables[0].Columns[10].ColumnName = "K8";
            dataSet.Tables[0].Columns[11].ColumnName = "L8";
            dataSet.Tables[0].Columns[12].ColumnName = "M8";
            dataSet.Tables[0].Columns[13].ColumnName = "N8";
            dataSet.Tables[0].Columns[14].ColumnName = "O8";
            dataSet.Tables[0].Columns[15].ColumnName = "P8";
            dataSet.Tables[0].Columns[16].ColumnName = "Q8";
            dataSet.Tables[0].Columns[17].ColumnName = "R8";
            dataSet.Tables[0].Columns[18].ColumnName = "S8";
            dataSet.Tables[0].Columns[19].ColumnName = "T8";
            dataSet.Tables[0].Columns[20].ColumnName = "U8";
            dataSet.Tables[0].Columns[21].ColumnName = "V8";
            dataSet.Tables[0].Columns[22].ColumnName = "W8";
            dataSet.Tables[0].Columns[23].ColumnName = "X8";
            dataSet.Tables[0].Columns[24].ColumnName = "Y8";
            //    dataSet.Tables[0].Columns[25].ColumnName = "Z8";

            return dataSet;
        }
        public DataSet DatTenEDO1_HNXS(DataSet dataSet)
        {
            //B8,C8,D8,E8,F8,G8,H8,I8,J8,K8,L8,M8,N8,O8,P8,Q8,R8,S8,T8,U8,V8,W8,X8,Y8,Z8
            dataSet.Tables[0].Columns[1].ColumnName = "B8";
            dataSet.Tables[0].Columns[2].ColumnName = "C8";
            dataSet.Tables[0].Columns[3].ColumnName = "D8";
            dataSet.Tables[0].Columns[4].ColumnName = "E8";
            dataSet.Tables[0].Columns[5].ColumnName = "F8";
            dataSet.Tables[0].Columns[6].ColumnName = "G8";
            dataSet.Tables[0].Columns[7].ColumnName = "H8";
            dataSet.Tables[0].Columns[8].ColumnName = "I8";
            dataSet.Tables[0].Columns[9].ColumnName = "J8";
            dataSet.Tables[0].Columns[10].ColumnName = "K8";
            dataSet.Tables[0].Columns[11].ColumnName = "L8";
            dataSet.Tables[0].Columns[12].ColumnName = "M8";
            dataSet.Tables[0].Columns[13].ColumnName = "N8";
            dataSet.Tables[0].Columns[14].ColumnName = "O8";
            dataSet.Tables[0].Columns[15].ColumnName = "P8";
            dataSet.Tables[0].Columns[16].ColumnName = "Q8";
            dataSet.Tables[0].Columns[17].ColumnName = "R8";
            dataSet.Tables[0].Columns[18].ColumnName = "S8";
            dataSet.Tables[0].Columns[19].ColumnName = "T8";
            dataSet.Tables[0].Columns[20].ColumnName = "U8";
            dataSet.Tables[0].Columns[21].ColumnName = "V8";
            dataSet.Tables[0].Columns[22].ColumnName = "W8";
            dataSet.Tables[0].Columns[23].ColumnName = "X8";
            dataSet.Tables[0].Columns[24].ColumnName = "Y8";
            dataSet.Tables[0].Columns[25].ColumnName = "Z8";
            dataSet.Tables[0].Columns[26].ColumnName = "AA8";
            dataSet.Tables[0].Columns[27].ColumnName = "AB8";

            return dataSet;
        }
        public DataSet DatTenNY_KQGD_Phien(DataSet dataSet)
        {
            //B8,C8,D8,E8,F8,G8,H8,I8,J8,K8,L8,M8,N8,O8,P8,Q8,R8,S8,T8,U8,V8,W8,X8,Y8,Z8
            dataSet.Tables[0].Columns[0].ColumnName = "A6";
            dataSet.Tables[0].Columns[1].ColumnName = "B6";
            dataSet.Tables[0].Columns[2].ColumnName = "C6";
            dataSet.Tables[0].Columns[3].ColumnName = "D6";
            dataSet.Tables[0].Columns[4].ColumnName = "E6";
            dataSet.Tables[0].Columns[5].ColumnName = "F6";
            dataSet.Tables[0].Columns[6].ColumnName = "G6";
            dataSet.Tables[0].Columns[7].ColumnName = "H6";
            dataSet.Tables[0].Columns[8].ColumnName = "I6";
            dataSet.Tables[0].Columns[9].ColumnName = "J6";
            dataSet.Tables[0].Columns[10].ColumnName = "K6";
            dataSet.Tables[0].Columns[11].ColumnName = "L6";
            dataSet.Tables[0].Columns[12].ColumnName = "M6";
            dataSet.Tables[0].Columns[13].ColumnName = "N6";
            dataSet.Tables[0].Columns[14].ColumnName = "O6";
            dataSet.Tables[0].Columns[15].ColumnName = "P6";
            dataSet.Tables[0].Columns[16].ColumnName = "Q6";
            dataSet.Tables[0].Columns[17].ColumnName = "R6";
            dataSet.Tables[0].Columns[18].ColumnName = "S6";
            dataSet.Tables[0].Columns[19].ColumnName = "T6";
            dataSet.Tables[0].Columns[20].ColumnName = "U6";
            dataSet.Tables[0].Columns[21].ColumnName = "V6";
            dataSet.Tables[0].Columns[22].ColumnName = "W6";


            return dataSet;
        }
        public DataSet DatTenNY_KQGD_Phien2(DataSet dataSet)
        {
            //B8,C8,D8,E8,F8,G8,H8,I8,J8,K8,L8,M8,N8,O8,P8,Q8,R8,S8,T8,U8,V8,W8,X8,Y8,Z8
            dataSet.Tables[0].Columns[0].ColumnName = "A6";
            dataSet.Tables[0].Columns[1].ColumnName = "B6";
            dataSet.Tables[0].Columns[2].ColumnName = "C6";
            dataSet.Tables[0].Columns[3].ColumnName = "D6";
            dataSet.Tables[0].Columns[4].ColumnName = "E6";
            dataSet.Tables[0].Columns[5].ColumnName = "F6";
            dataSet.Tables[0].Columns[6].ColumnName = "G6";
            dataSet.Tables[0].Columns[7].ColumnName = "H6";
            dataSet.Tables[0].Columns[8].ColumnName = "I6";
            dataSet.Tables[0].Columns[9].ColumnName = "J6";
            dataSet.Tables[0].Columns[10].ColumnName = "K6";
            dataSet.Tables[0].Columns[11].ColumnName = "L6";
            dataSet.Tables[0].Columns[12].ColumnName = "M6";
            dataSet.Tables[0].Columns[13].ColumnName = "N6";
            dataSet.Tables[0].Columns[14].ColumnName = "O6";
            dataSet.Tables[0].Columns[15].ColumnName = "P6";
            dataSet.Tables[0].Columns[16].ColumnName = "Q6";
            dataSet.Tables[0].Columns[17].ColumnName = "R6";
            dataSet.Tables[0].Columns[18].ColumnName = "S6";
            dataSet.Tables[0].Columns[19].ColumnName = "T6";
            dataSet.Tables[0].Columns[20].ColumnName = "U6";
            dataSet.Tables[0].Columns[21].ColumnName = "V6";
            // dataSet.Tables[0].Columns[22].ColumnName = "W6";


            return dataSet;
        }

        public DataSet DatTenNY_KQGD_Phien2SS(DataSet dataSet)
        {
            //B8,C8,D8,E8,F8,G8,H8,I8,J8,K8,L8,M8,N8,O8,P8,Q8,R8,S8,T8,U8,V8,W8,X8,Y8,Z8
            dataSet.Tables[0].Columns[0].ColumnName = "A6";
            dataSet.Tables[0].Columns[1].ColumnName = "B6";
            dataSet.Tables[0].Columns[2].ColumnName = "C6";
            dataSet.Tables[0].Columns[3].ColumnName = "D6";
            dataSet.Tables[0].Columns[4].ColumnName = "E6";
            dataSet.Tables[0].Columns[5].ColumnName = "F6";
            dataSet.Tables[0].Columns[6].ColumnName = "G6";
            dataSet.Tables[0].Columns[7].ColumnName = "H6";
            dataSet.Tables[0].Columns[8].ColumnName = "I6";
            dataSet.Tables[0].Columns[9].ColumnName = "J6";
            dataSet.Tables[0].Columns[10].ColumnName = "K6";
            dataSet.Tables[0].Columns[11].ColumnName = "L6";
            dataSet.Tables[0].Columns[12].ColumnName = "M6";
            dataSet.Tables[0].Columns[13].ColumnName = "N6";
            dataSet.Tables[0].Columns[14].ColumnName = "O6";
            dataSet.Tables[0].Columns[15].ColumnName = "P6";
            dataSet.Tables[0].Columns[16].ColumnName = "Q6";
            dataSet.Tables[0].Columns[17].ColumnName = "R6";
            dataSet.Tables[0].Columns[18].ColumnName = "S6";
            dataSet.Tables[0].Columns[19].ColumnName = "T6";
            dataSet.Tables[0].Columns[20].ColumnName = "U6";
            //  dataSet.Tables[0].Columns[21].ColumnName = "V6";
            // dataSet.Tables[0].Columns[22].ColumnName = "W6";


            return dataSet;
        }


        public DataSet DatTenEDO7(DataSet dataSet)
        {
            //B8,C8,D8,E8,F8,G8,H8,I8,J8,K8,L8,M8,N8,O8,P8,Q8,R8,S8,T8,U8,V8,W8,X8,Y8,Z8
            dataSet.Tables[0].Columns[1].ColumnName = "B8";
            dataSet.Tables[0].Columns[2].ColumnName = "C8";
            dataSet.Tables[0].Columns[3].ColumnName = "D8";
            dataSet.Tables[0].Columns[4].ColumnName = "E8";
            dataSet.Tables[0].Columns[5].ColumnName = "F8";
            dataSet.Tables[0].Columns[6].ColumnName = "G8";
            dataSet.Tables[0].Columns[7].ColumnName = "H8";
            dataSet.Tables[0].Columns[8].ColumnName = "I8";
            dataSet.Tables[0].Columns[9].ColumnName = "J8";

            return dataSet;
        }
        public DataSet DatTenEDO2_1(DataSet dataSet)
        {
            //B4,C4,D4
            dataSet.Tables[0].Columns[1].ColumnName = "B4";
            dataSet.Tables[0].Columns[2].ColumnName = "C4";
            dataSet.Tables[0].Columns[3].ColumnName = "D4";

            return dataSet;
        }
        public DataSet DatTenUPCOMEDO2_1(DataSet dataSet)
        {
            //B4,C4,D4
            dataSet.Tables[0].Columns[0].ColumnName = "A4";
            dataSet.Tables[0].Columns[1].ColumnName = "B4";
            //   dataSet.Tables[0].Columns[3].ColumnName = "D4";

            return dataSet;
        }
        public DataSet DatTenNY22_1(DataSet dataSet)
        {
            //B4,C4,D4
            dataSet.Tables[0].Columns[0].ColumnName = "A4";
            dataSet.Tables[0].Columns[1].ColumnName = "B4";


            return dataSet;
        }
        public DataSet DatTenNY22_2(DataSet dataSet)
        {
            //E5,F5,G5,H5
            dataSet.Tables[0].Columns[4].ColumnName = "E5";
            dataSet.Tables[0].Columns[5].ColumnName = "F5";
            dataSet.Tables[0].Columns[6].ColumnName = "G5";
            dataSet.Tables[0].Columns[7].ColumnName = "H5";
            //dataSet.Tables[0].Columns[8].ColumnName = "I6";
            return dataSet;
        }

        public DataSet DatTenEDO2_2(DataSet dataSet)
        {
            //F6,G6,H6
            dataSet.Tables[0].Columns[5].ColumnName = "F6";
            dataSet.Tables[0].Columns[6].ColumnName = "G6";
            dataSet.Tables[0].Columns[7].ColumnName = "H6";
            //dataSet.Tables[0].Columns[8].ColumnName = "I6";
            return dataSet;
        }
        //E5,F5,G5
        public DataSet DatTenUPCOMEOD02_2(DataSet dataSet)
        {
            //F6,G6,H6
            dataSet.Tables[0].Columns[4].ColumnName = "E5";
            dataSet.Tables[0].Columns[5].ColumnName = "F5";
            dataSet.Tables[0].Columns[6].ColumnName = "G5";
            //dataSet.Tables[0].Columns[8].ColumnName = "I6";
            return dataSet;
        }
        public DataSet DatTenEDO2_2_HNX(DataSet dataSet)
        {
            //F6,G6,H6
            dataSet.Tables[0].Columns[5].ColumnName = "F5";
            dataSet.Tables[0].Columns[6].ColumnName = "G5";
            dataSet.Tables[0].Columns[7].ColumnName = "H5";
            dataSet.Tables[0].Columns[8].ColumnName = "I5";
            return dataSet;
        }
        public DataSet DatTenEDO2_3(DataSet dataSet)
        {
            //F21,G21,H21,I21
            dataSet.Tables[0].Columns[5].ColumnName = "F21";
            dataSet.Tables[0].Columns[6].ColumnName = "G21";
            dataSet.Tables[0].Columns[7].ColumnName = "H21";
            dataSet.Tables[0].Columns[8].ColumnName = "I21";
            return dataSet;
        }
        public DataSet DatTenUPCOMEDO2_3(DataSet dataSet)
        {
            //E23,F23,G23,H23
            dataSet.Tables[0].Columns[4].ColumnName = "E23";
            dataSet.Tables[0].Columns[5].ColumnName = "F23";
            dataSet.Tables[0].Columns[6].ColumnName = "G23";
            dataSet.Tables[0].Columns[7].ColumnName = "H23";

            return dataSet;
        }

        public DataSet DatTenNY22_3(DataSet dataSet)
        {
            //F21,G21,H21,I21
            dataSet.Tables[0].Columns[4].ColumnName = "E21";
            dataSet.Tables[0].Columns[5].ColumnName = "F21";
            dataSet.Tables[0].Columns[6].ColumnName = "G21";
            dataSet.Tables[0].Columns[7].ColumnName = "H21";

            return dataSet;
        }
        public DataSet DatTenEDO2_4_HNX(DataSet dataSet)
        {
            //F21,G21,H21,I21
            dataSet.Tables[0].Columns[5].ColumnName = "F36";
            dataSet.Tables[0].Columns[6].ColumnName = "G36";
            dataSet.Tables[0].Columns[7].ColumnName = "H36";
            dataSet.Tables[0].Columns[8].ColumnName = "I36";
            return dataSet;
        }
        public DataSet DatTenNY22_4(DataSet dataSet)
        {
            //F21,G21,H21,I21
            dataSet.Tables[0].Columns[4].ColumnName = "E35";
            dataSet.Tables[0].Columns[5].ColumnName = "F35";
            dataSet.Tables[0].Columns[6].ColumnName = "G35";
            dataSet.Tables[0].Columns[7].ColumnName = "H35";

            return dataSet;
        }
        public DataSet DatTenEDO2_5_HNX(DataSet dataSet)
        {
            //F21,G21,H21,I21
            dataSet.Tables[0].Columns[5].ColumnName = "F51";
            dataSet.Tables[0].Columns[6].ColumnName = "G51";
            dataSet.Tables[0].Columns[7].ColumnName = "H51";
            dataSet.Tables[0].Columns[8].ColumnName = "I51";
            dataSet.Tables[0].Columns[9].ColumnName = "J51";
            return dataSet;
        }
        public DataSet DatTenNY22_5(DataSet dataSet)
        {
            //F21,G21,H21,I21
            dataSet.Tables[0].Columns[4].ColumnName = "E49";
            dataSet.Tables[0].Columns[5].ColumnName = "F49";
            dataSet.Tables[0].Columns[6].ColumnName = "G49";
            dataSet.Tables[0].Columns[7].ColumnName = "H49";
            dataSet.Tables[0].Columns[8].ColumnName = "I49";

            return dataSet;
        }
        public DataSet DatTenEDO2_4(DataSet dataSet)
        {
            //F6,G6,H6
            dataSet.Tables[0].Columns[11].ColumnName = "L6";
            dataSet.Tables[0].Columns[12].ColumnName = "M6";
            dataSet.Tables[0].Columns[13].ColumnName = "N6";

            return dataSet;
        }
        public DataSet DatTenUPCOMEDO2_4(DataSet dataSet)
        {
            //I5,J5,K5
            dataSet.Tables[0].Columns[8].ColumnName = "I5";
            dataSet.Tables[0].Columns[9].ColumnName = "J5";
            dataSet.Tables[0].Columns[10].ColumnName = "K5";

            return dataSet;
        }
        public DataSet DatTenEDO2_6_HNX(DataSet dataSet)
        {
            //F6,G6,H6
            dataSet.Tables[0].Columns[11].ColumnName = "L5";
            dataSet.Tables[0].Columns[12].ColumnName = "M5";
            dataSet.Tables[0].Columns[13].ColumnName = "N5";
            dataSet.Tables[0].Columns[14].ColumnName = "O5";
            return dataSet;
        }
        public DataSet DatTenNY22_6(DataSet dataSet)
        {
            //F6,G6,H6
            dataSet.Tables[0].Columns[9].ColumnName = "J5";
            dataSet.Tables[0].Columns[10].ColumnName = "K5";
            dataSet.Tables[0].Columns[11].ColumnName = "L5";
            dataSet.Tables[0].Columns[12].ColumnName = "M5";

            return dataSet;
        }
        public DataSet DatTenEDO2_7_HNX(DataSet dataSet)
        {
            //F6,G6,H6
            dataSet.Tables[0].Columns[11].ColumnName = "L21";
            dataSet.Tables[0].Columns[12].ColumnName = "M21";
            dataSet.Tables[0].Columns[13].ColumnName = "N21";
            dataSet.Tables[0].Columns[14].ColumnName = "O21";
            return dataSet;
        }
        public DataSet DatTenNY22_7(DataSet dataSet)
        {
            //F6,G6,H6
            dataSet.Tables[0].Columns[9].ColumnName = "J21";
            dataSet.Tables[0].Columns[10].ColumnName = "K21";
            dataSet.Tables[0].Columns[11].ColumnName = "L21";
            dataSet.Tables[0].Columns[12].ColumnName = "M21";

            return dataSet;
        }
        public DataSet DatTenEDO2_8_HNX(DataSet dataSet)
        {
            //F6,G6,H6
            dataSet.Tables[0].Columns[11].ColumnName = "L36";
            dataSet.Tables[0].Columns[12].ColumnName = "M36";
            dataSet.Tables[0].Columns[13].ColumnName = "N36";
            dataSet.Tables[0].Columns[14].ColumnName = "O36";
            return dataSet;
        }
        public DataSet DatTenNY22_8(DataSet dataSet)
        {
            //F6,G6,H6
            dataSet.Tables[0].Columns[9].ColumnName = "J35";
            dataSet.Tables[0].Columns[10].ColumnName = "K35";
            dataSet.Tables[0].Columns[11].ColumnName = "L35";
            dataSet.Tables[0].Columns[12].ColumnName = "M35";

            return dataSet;
        }
        public DataSet DatTenEDO2_9_HNX(DataSet dataSet)
        {
            //F6,G6,H6
            dataSet.Tables[0].Columns[11].ColumnName = "L51";
            dataSet.Tables[0].Columns[12].ColumnName = "M51";
            dataSet.Tables[0].Columns[13].ColumnName = "N51";
            dataSet.Tables[0].Columns[14].ColumnName = "O51";
            dataSet.Tables[0].Columns[15].ColumnName = "P51";
            return dataSet;
        }
        public DataSet DatTenNY22_9(DataSet dataSet)
        {
            //F6,G6,H6
            dataSet.Tables[0].Columns[10].ColumnName = "K49";
            dataSet.Tables[0].Columns[11].ColumnName = "L49";
            dataSet.Tables[0].Columns[12].ColumnName = "M49";
            dataSet.Tables[0].Columns[13].ColumnName = "N49";
            dataSet.Tables[0].Columns[14].ColumnName = "O49";

            return dataSet;
        }
        public DataSet DatTenEDO2_5(DataSet dataSet)
        {
            //F21,G21,H21,I21
            dataSet.Tables[0].Columns[11].ColumnName = "L21";
            dataSet.Tables[0].Columns[12].ColumnName = "M21";
            dataSet.Tables[0].Columns[13].ColumnName = "N21";
            dataSet.Tables[0].Columns[14].ColumnName = "O21";
            return dataSet;
        }

        public DataSet DatTenUPCOMEDO2_5(DataSet dataSet)
        {
            //J23,K23,L23,M23
            dataSet.Tables[0].Columns[9].ColumnName = "J23";
            dataSet.Tables[0].Columns[10].ColumnName = "K23";
            dataSet.Tables[0].Columns[11].ColumnName = "L23";
            dataSet.Tables[0].Columns[12].ColumnName = "M23";

            return dataSet;
        }


        public EDalResult Update_2017(NY_KQGD kqgd, NY_ThongKeCC cc_hnx, NY_GDDTNN gddtnn, NY_TTCP ttcp, bool getScriptOnly = false)
        {

            try
            {
                string spName = "";
                EDalResult result;
                //   DynamicParameters dynamicParameters = new DynamicParameters();
                StringBuilder sb = new StringBuilder();
                //STT,Symbol,BasicPrice,OpenPrice,ClosePrice,HighestPrice,LowestPrice,
                //Diem,PhanTram,KLGD_KL,GTGD_KL,KLGD_TT,GTGD_TT,KLGD_LL,GTGD_LL,KLGD_TC
                //,TyTrong1,GTGD_TC,TyTrong2,KLCP_LuuHanh,GTVHTT_GT,GTVHTT_TT,VDL,Trangding_Date
                if (kqgd != null)
                {
                    sb.Append("(");
                    sb.Append(kqgd.STT).Append(",");
                    sb.Append("'" + kqgd.Symbol + "'").Append(",");
                    sb.Append(kqgd.BasicPrice).Append(",");
                    sb.Append(kqgd.OpenPrice).Append(",");
                    sb.Append(kqgd.ClosePrice).Append(",");
                    sb.Append(kqgd.HighestPrice).Append(",");
                    sb.Append(kqgd.LowestPrice).Append(",");

                    sb.Append(kqgd.Diem).Append(",");

                    sb.Append(kqgd.PhanTram).Append(",");
                    sb.Append(kqgd.KLGD_KL).Append(",");
                    sb.Append(kqgd.GTGD_KL).Append(",");

                    sb.Append(kqgd.KLGD_TT).Append(",");
                    sb.Append(kqgd.GTGD_TT).Append(",");

                    sb.Append(kqgd.KLGD_LL).Append(",");
                    sb.Append(kqgd.GTGD_LL).Append(",");
                    sb.Append(kqgd.KLGD_TC).Append(",");
                    sb.Append(kqgd.TyTrong1).Append(",");
                    sb.Append(kqgd.GTGD_TC).Append(",");
                    sb.Append(kqgd.TyTrong2).Append(",");

                    sb.Append(kqgd.KLCP_LuuHanh).Append(",");
                    sb.Append(kqgd.GTVHTT_GT).Append(",");
                    sb.Append(kqgd.GTVHTT_TT).Append(",");
                    sb.Append(kqgd.VDL).Append(",");
                    sb.Append("'" + kqgd.Trangding_Date + "'");
                    sb.Append("),");
                }
                if (cc_hnx != null)
                {
                    sb.Append("(");
                    sb.Append(cc_hnx.STT).Append(",");
                    sb.Append("'" + cc_hnx.Symbol + "'").Append(",");
                    sb.Append(cc_hnx.SLDATMUA_KL).Append(",");
                    sb.Append(cc_hnx.KLDATMUA_KL).Append(",");
                    sb.Append(cc_hnx.SLDATBAN_KL).Append(",");
                    sb.Append(cc_hnx.KLDATBAN_KL).Append(",");
                    sb.Append(cc_hnx.SLDATMUA_TT).Append(",");
                    sb.Append(cc_hnx.KLDATMUA_TT).Append(",");
                    sb.Append(cc_hnx.SLDATBAN_TT).Append(",");
                    sb.Append(cc_hnx.KLDATBAN_TT).Append(",");
                    sb.Append(cc_hnx.SLDATMUA_TC).Append(",");
                    sb.Append(cc_hnx.KLDATMUA_TC).Append(",");
                    sb.Append(cc_hnx.SLDATBAN_TC).Append(",");
                    sb.Append(cc_hnx.KLDATBAN_TC).Append(",");
                    sb.Append(cc_hnx.KLDUMUA).Append(",");
                    sb.Append(cc_hnx.KLDUBAN).Append(",");
                    sb.Append(cc_hnx.KLTHUCHIEN).Append(",");
                    sb.Append(cc_hnx.GTTHUCHIEN).Append(",");
                    sb.Append("'" + cc_hnx.Trangding_Date + "'");
                    sb.Append("),");

                }
                //STT,Symbol,KLCP_NY,KLCP_LH,Co_Tuc_2014,Co_Tuc_2015,PE,EPS2015,
                //ROE2015,ROA2015,BasicPrice_KT,CeilingPrice_KT,FloorPrice_KT,Trangding_Date
                if (ttcp != null)
                {
                    sb.Append("(");
                    sb.Append(ttcp.STT).Append(",");
                    sb.Append("'" + ttcp.Symbol + "'").Append(",");
                    sb.Append(ttcp.KLCP_NY).Append(",");
                    sb.Append(ttcp.KLCP_LH).Append(",");
                    sb.Append(ttcp.Co_Tuc_2014).Append(",");
                    sb.Append(ttcp.Co_Tuc_2015).Append(",");
                    sb.Append(ttcp.PE).Append(",");
                    sb.Append(ttcp.EPS2015).Append(",");
                    sb.Append(ttcp.ROE2015).Append(",");
                    sb.Append(ttcp.ROA2015).Append(",");
                    sb.Append(ttcp.BasicPrice_KT).Append(",");
                    sb.Append(ttcp.CeilingPrice_KT).Append(",");
                    sb.Append(ttcp.FloorPrice_KT).Append(",");
                    sb.Append("'" + ttcp.Trangding_Date + "'");
                    sb.Append("),");

                }
                if (gddtnn != null)
                {
                    sb.Append("(");
                    sb.Append(gddtnn.STT).Append(",");
                    sb.Append("'" + gddtnn.Symbol + "'").Append(",");
                    sb.Append(gddtnn.KLMUA_KL).Append(",");
                    sb.Append(gddtnn.GTMUA_KL).Append(",");
                    sb.Append(gddtnn.KLBAN_KL).Append(",");
                    sb.Append(gddtnn.GTBAN_KL).Append(",");
                    sb.Append(gddtnn.KLMUA_TT).Append(",");

                    sb.Append(gddtnn.GTMUA_TT).Append(",");

                    sb.Append(gddtnn.KLBAN_TT).Append(",");
                    sb.Append(gddtnn.GTBAN_TT).Append(",");
                    sb.Append(gddtnn.KLMUA_TC).Append(",");
                    sb.Append(gddtnn.GTMUA_TC).Append(",");
                    sb.Append(gddtnn.KLBAN_TC).Append(",");
                    sb.Append(gddtnn.GTBAN_TC).Append(",");
                    sb.Append(gddtnn.KLCK_MAX).Append(",");
                    sb.Append(gddtnn.KLCK_NDTNN).Append(",");
                    sb.Append(gddtnn.KLCK_CDPNG).Append(",");
                    sb.Append("'" + gddtnn.Trangding_Date + "'");
                    sb.Append("),");

                }

                // ko exec sp, chi lay script de run bulk update sau nay 
                /* if (getScriptOnly)
                 {*/
                return new EDalResult() { Code = EDalResult.__CODE_SUCCESS, Message = EDalResult.__STRING_GET_SCRIPT, Data = sb.ToString() };
                // }

                /*  // 2. main			
                  result = await ExecuteSpNoQueryAsync(spName, dynamicParameters);


                  // return (neu sp ko tra error code,msg thi tu gan default)
                  return new EDalResult() { Code = EDalResult.__CODE_SUCCESS, Message = EDalResult.__STRING_SUCCESS, Data = result.Data };*/
            }
            catch (Exception ex)
            {
                // log error + buffer data
                //  this._cS6GApp.ErrorLogger.LogErrorContext(ex, ec);
                // error => return null
                return new EDalResult() { Code = -9997, Message = ex.Message, Data = null };
            }
        }

        //2011

        public EDalResult Update_TTICBVCTDKGD_HNX_2011(KQGIAODICHCP2011 kq_hnx, TinhHinhDatLenh2011 thdl, NDTNN2011 ndtnn, KQGDCHITIET2011 kqct, KQGDTH2011 kqgdth, Top10CK_GTGDL gtgdl
            , Top10CK_KLGDL klgdl, Top10CP_GTNYL ctnyl, Top10CK_TANGGIA tanggia, Top10CK_GIAMGIA giamgia, Chi_Tieu_2011 ct, Top10CK_NDTNN ndtnns, KLGD_TOP2011_MR mr1,
           GTGD_TOP2011_MR mr2, TangGiam_TOP2011_MR mr3, CKNTDNN_TOP2011_MR mr4, bool getScriptOnly = false)
        {

            try
            {
                string spName = "";
                EDalResult result;
                //   DynamicParameters dynamicParameters = new DynamicParameters();
                StringBuilder sb = new StringBuilder();
                if (mr4 != null)
                {
                    //Symbol,KLMua,GTMua,KLDPNamGiu,Trangding_Date
                    sb.Append("(");

                    sb.Append("'" + mr4.Symbol + "'").Append(",");
                    sb.Append(mr4.KLMua).Append(",");
                    sb.Append(mr4.GTMua).Append(",");
                    sb.Append(mr4.KLBan).Append(",");
                    sb.Append(mr4.GTBan).Append(",");
                    sb.Append(mr4.KLDPNamGiu).Append(",");
                    sb.Append("'" + mr4.Trangding_Date + "'");
                    sb.Append("),");

                }
                if (mr3 != null)
                {
                    //Symbol,AvePrice,MucTang,PTTangGiam,KLGD,CEILINGPRICE,ChenhLechTran,FLOORPRICES,ChenhLechSan,Trangding_Date
                    sb.Append("(");

                    sb.Append("'" + mr3.Symbol + "'").Append(",");
                    sb.Append(mr3.AvePrice).Append(",");
                    sb.Append(mr3.MucTang).Append(",");
                    sb.Append(mr3.PTTangGiam).Append(",");
                    sb.Append(mr3.KLGD).Append(",");
                    sb.Append(mr3.CEILINGPRICE).Append(",");
                    sb.Append(mr3.ChenhLechTran).Append(",");
                    sb.Append(mr3.FLOORPRICES).Append(",");
                    sb.Append(mr3.ChenhLechSan).Append(",");
                    sb.Append("'" + mr3.Trangding_Date + "'");
                    sb.Append("),");

                }
                if (mr2 != null)
                {
                    //Symbol,AvePrice,KLGD,KLNY,GTNY_Trieu,GTNY_Dong,Trangding_Date
                    sb.Append("(");

                    sb.Append("'" + mr2.Symbol + "'").Append(",");
                    sb.Append(mr2.AvePrice).Append(",");
                    sb.Append(mr2.KLGD).Append(",");
                    sb.Append(mr2.KLNY).Append(",");
                    sb.Append(mr2.GTNY_Trieu).Append(",");
                    sb.Append(mr2.GTNY_Dong).Append(",");
                    sb.Append("'" + mr2.Trangding_Date + "'");
                    sb.Append("),");

                }
                if (mr1 != null)
                {
                    //Symbol,AvePrice,KL,GT,TangGiam,KLGD_NgayTruoc,Trangding_Date
                    sb.Append("(");

                    sb.Append("'" + mr1.Symbol + "'").Append(",");
                    sb.Append(mr1.AvePrice).Append(",");
                    sb.Append(mr1.KL).Append(",");
                    sb.Append(mr1.GT).Append(",");
                    sb.Append(mr1.TangGiam).Append(",");
                    sb.Append(mr1.KLGD_NgayTruoc).Append(",");
                    sb.Append("'" + mr1.Trangding_Date + "'");
                    sb.Append("),");

                }
                if (ndtnns != null)
                {
                    //Symbol,KLMua,GTMua,KLDPNamGiu,Trangding_Date
                    sb.Append("(");

                    sb.Append("'" + ndtnns.Symbol + "'").Append(",");
                    sb.Append(ndtnns.KLMua).Append(",");
                    sb.Append(ndtnns.GTMua).Append(",");
                    sb.Append(ndtnns.KLDPNamGiu).Append(",");
                    sb.Append("'" + ndtnns.Trangding_Date + "'");
                    sb.Append("),");

                }
                if (ct != null)
                {
                    //Chi_Tieu,CPNY,CP_DKGD_UPCOM,Trangding_Date
                    sb.Append("(");

                    sb.Append("'" + ct.Chi_Tieu + "'").Append(",");
                    sb.Append(ct.CPNY).Append(",");
                    sb.Append(ct.CP_DKGD_UPCOM).Append(",");

                    sb.Append("'" + ct.Trangding_Date + "'");
                    sb.Append("),");

                }
                if (giamgia != null)
                {
                    //Symbol,AvePrice,MucGiam,TyLeTang,Trangding_Date
                    sb.Append("(");

                    sb.Append("'" + giamgia.Symbol + "'").Append(",");
                    sb.Append(giamgia.AvePrice).Append(",");
                    sb.Append(giamgia.MucGiam).Append(",");
                    sb.Append(giamgia.TyLeTang).Append(",");

                    sb.Append("'" + giamgia.Trangding_Date + "'");
                    sb.Append("),");

                }
                if (tanggia != null)
                {
                    //Symbol,AvePrice,TyLeTang,KLGD,Trangding_Date
                    sb.Append("(");

                    sb.Append("'" + tanggia.Symbol + "'").Append(",");
                    sb.Append(tanggia.AvePrice).Append(",");
                    sb.Append(tanggia.TyLeTang).Append(",");
                    sb.Append(tanggia.KLGD).Append(",");

                    sb.Append("'" + tanggia.Trangding_Date + "'");
                    sb.Append("),");

                }
                if (ctnyl != null)
                {
                    //Symbol,AvePrice,Volume,GiaTriNY,Trangding_Date
                    sb.Append("(");

                    sb.Append("'" + ctnyl.Symbol + "'").Append(",");
                    sb.Append(ctnyl.AvePrice).Append(",");
                    sb.Append(ctnyl.Volume).Append(",");
                    sb.Append(ctnyl.GiaTriNY).Append(",");

                    sb.Append("'" + ctnyl.Trangding_Date + "'");
                    sb.Append("),");

                }
                if (klgdl != null)
                {
                    // Symbol,AvePrice,Volume,PhanTram,WeightN,Trangding_Date
                    sb.Append("(");

                    sb.Append("'" + klgdl.Symbol + "'").Append(",");
                    sb.Append(klgdl.AvePrice).Append(",");
                    sb.Append(klgdl.Volume).Append(",");
                    sb.Append(klgdl.PhanTram).Append(",");
                    sb.Append(klgdl.WeightN).Append(",");

                    sb.Append("'" + klgdl.Trangding_Date + "'");
                    sb.Append("),");

                }
                if (kq_hnx != null)
                {
                    // //STT,Symbol,SLCP_DKGD,SLCP_LH,Co_Tuc_2010,PE,EPS2010,KLGD_10PHIEN,
                    //ROE,ROA,BasicPrice_KT,CeilingPrice_KT,FloorPrice_KT,Co_Tuc_2009,Trangding_Date
                    sb.Append("(");
                    sb.Append(kq_hnx.STT).Append(",");
                    sb.Append("'" + kq_hnx.Symbol + "'").Append(",");
                    sb.Append(kq_hnx.SLCP_DKGD).Append(",");
                    sb.Append(kq_hnx.SLCP_LH).Append(",");
                    sb.Append(kq_hnx.Co_Tuc_2010).Append(",");
                    sb.Append(kq_hnx.PE).Append(",");
                    sb.Append(kq_hnx.EPS2010).Append(",");

                    sb.Append(kq_hnx.KLGD_10PHIEN).Append(",");

                    sb.Append(kq_hnx.ROE).Append(",");
                    sb.Append(kq_hnx.ROA).Append(",");
                    sb.Append(kq_hnx.BasicPrice_KT).Append(",");
                    sb.Append(kq_hnx.CeilingPrice_KT).Append(",");
                    sb.Append(kq_hnx.FloorPrice_KT).Append(",");
                    sb.Append(kq_hnx.BinhQuan).Append(",");
                    sb.Append(kq_hnx.Tong).Append(",");
                    sb.Append(kq_hnx.Co_Tuc_2009).Append(",");
                    sb.Append("'" + kq_hnx.Trangding_Date + "'");
                    sb.Append("),");

                }
                if (gtgdl != null)
                {
                    // Symbol,ValueN,WeightN,Trangding_Date
                    sb.Append("(");

                    sb.Append("'" + gtgdl.Symbol + "'").Append(",");
                    sb.Append(gtgdl.ValueN).Append(",");
                    sb.Append(gtgdl.WeightN).Append(",");
                    sb.Append("'" + gtgdl.Trangding_Date + "'");
                    sb.Append("),");

                }
                if (kqgdth != null)
                {
                    // TypeName,Volume_BG,Value_BG,Weight_BG,Volume_TT,Value_TT,Weight_TT,Volume_MT,Value_MT,Weight_MT,Trangding_Date
                    sb.Append("(");

                    sb.Append("'" + kqgdth.TypeName + "'").Append(",");
                    sb.Append(kqgdth.Volume_BG).Append(",");
                    sb.Append(kqgdth.Value_BG).Append(",");
                    sb.Append(kqgdth.Weight_BG).Append(",");
                    sb.Append(kqgdth.Volume_TT).Append(",");
                    sb.Append(kqgdth.Value_TT).Append(",");

                    sb.Append(kqgdth.Weight_TT).Append(",");

                    sb.Append(kqgdth.Volume_MT).Append(",");
                    sb.Append(kqgdth.Value_MT).Append(",");
                    sb.Append(kqgdth.Weight_MT).Append(",");

                    sb.Append("'" + kqgdth.Trangding_Date + "'");
                    sb.Append("),");

                }
                if (kqct != null)
                {
                    // STT,Symbol,BasicPrice,OpenPrice,ClosePrice,HighPrice,LowPrice,AveragePrice,NetChange,Volume_BG,
                    // Value_BG,AveragePrice_TT,Volume_TT,Value_TT,Volume_TC,Value_TC,GiaTriTT,Trangding_Date
                    sb.Append("(");
                    sb.Append(kqct.STT).Append(",");
                    sb.Append("'" + kqct.Symbol + "'").Append(",");
                    sb.Append(kqct.BasicPrice).Append(",");
                    sb.Append(kqct.OpenPrice).Append(",");
                    sb.Append(kqct.ClosePrice).Append(",");
                    sb.Append(kqct.HighPrice).Append(",");
                    sb.Append(kqct.LowPrice).Append(",");
                    sb.Append(kqct.AveragePrice).Append(",");
                    sb.Append(kqct.NetChange).Append(",");
                    sb.Append(kqct.Volume_BG).Append(",");
                    sb.Append(kqct.Value_BG).Append(",");
                    sb.Append(kqct.AveragePrice_TT).Append(",");
                    sb.Append(kqct.Volume_TT).Append(",");
                    sb.Append(kqct.Value_TT).Append(",");
                    sb.Append(kqct.Volume_TC).Append(",");
                    sb.Append(kqct.Value_TC).Append(",");
                    sb.Append(kqct.GiaTriTT).Append(",");
                    sb.Append("'" + kqct.Trangding_Date + "'");
                    sb.Append("),");

                }
                if (ndtnn != null)
                {
                    //STT,Symbol,KLCKMAX,KLMUA_QT,GTMUA_QT,KLBAN_QT,GIATRI_QT,KLMUA_NT,GTMUA_NT,
                    //KLBAN_NT,GIATRI_NT,CurrentRoom,KLLH,NamGiuMax,KLNDTN,Trangding_Date
                    sb.Append("(");
                    sb.Append(ndtnn.STT).Append(",");
                    sb.Append("'" + ndtnn.Symbol + "'").Append(",");
                    sb.Append(ndtnn.KLCKMAX).Append(",");
                    sb.Append(ndtnn.KLMUA_QT).Append(",");
                    sb.Append(ndtnn.GTMUA_QT).Append(",");
                    sb.Append(ndtnn.KLBAN_QT).Append(",");
                    sb.Append(ndtnn.GIATRI_QT).Append(",");
                    sb.Append(ndtnn.KLMUA_NT).Append(",");
                    sb.Append(ndtnn.GTMUA_NT).Append(",");
                    sb.Append(ndtnn.KLBAN_NT).Append(",");
                    sb.Append(ndtnn.GIATRI_NT).Append(",");
                    sb.Append(ndtnn.CurrentRoom).Append(",");
                    sb.Append(ndtnn.KLLH).Append(",");
                    sb.Append(ndtnn.NamGiuMax).Append(",");
                    sb.Append(ndtnn.KLNDTN).Append(",");
                    sb.Append("'" + ndtnn.Trangding_Date + "'");
                    sb.Append("),");

                }
                if (thdl != null)
                {
                    //Symbol,NumberofBids_QT,BidVolume_QT,NumberofOffers_QT,OfferVolume_QT,Difference_QT,
                    //NumberofBids_NT,BidVolume_NT,NumberofOffers_NT,OfferVolume_NT,Difference_NT,SLDatMua,KLDatMua,SLDatBan,KLDatBan,Trangding_Date
                    sb.Append("(");
                    sb.Append("'" + thdl.Symbol + "'").Append(",");
                    sb.Append(thdl.NumberofBids_QT).Append(",");
                    sb.Append(thdl.BidVolume_QT).Append(",");
                    sb.Append(thdl.NumberofOffers_QT).Append(",");
                    sb.Append(thdl.OfferVolume_QT).Append(",");
                    sb.Append(thdl.Difference_QT).Append(",");
                    sb.Append(thdl.NumberofBids_NT).Append(",");
                    sb.Append(thdl.BidVolume_NT).Append(",");
                    sb.Append(thdl.NumberofOffers_NT).Append(",");
                    sb.Append(thdl.OfferVolume_NT).Append(",");
                    sb.Append(thdl.Difference_NT).Append(",");
                    sb.Append(thdl.SLDatMua).Append(",");
                    sb.Append(thdl.KLDatMua).Append(",");
                    sb.Append(thdl.SLDatBan).Append(",");
                    sb.Append(thdl.KLDatBan).Append(",");
                    sb.Append("'" + thdl.Trangding_Date + "'");
                    sb.Append("),");

                }

                // ko exec sp, chi lay script de run bulk update sau nay 
                /* if (getScriptOnly)
                 {*/
                return new EDalResult() { Code = EDalResult.__CODE_SUCCESS, Message = EDalResult.__STRING_GET_SCRIPT, Data = sb.ToString() };
                // }

                /*  // 2. main			
                  result = await ExecuteSpNoQueryAsync(spName, dynamicParameters);


                  // return (neu sp ko tra error code,msg thi tu gan default)
                  return new EDalResult() { Code = EDalResult.__CODE_SUCCESS, Message = EDalResult.__STRING_SUCCESS, Data = result.Data };*/
            }
            catch (Exception ex)
            {
                // log error + buffer data
                //  this._cS6GApp.ErrorLogger.LogErrorContext(ex, ec);
                // error => return null
                return new EDalResult() { Code = -9997, Message = ex.Message, Data = null };
            }
        }
        //2010
        public EDalResult Update_TTICBVCTDKGD_HNX_2010(Top10CK_TANGGIA2010 tanggia, Top10CP_CLGMAX max, GD_TRAIPHIEU tp, GDTP_NDTNN tc,
            bool getScriptOnly = false)
        {

            try
            {
                string spName = "";
                EDalResult result;
                StringBuilder sb = new StringBuilder();


                if (tanggia != null)
                {
                    //Symbol,AvePrice,TyLeTang,KLGD,Trangding_Date
                    sb.Append("(");
                    sb.Append("'" + tanggia.Symbol + "'").Append(",");
                    sb.Append(tanggia.AvePrice).Append(",");
                    sb.Append(tanggia.MucTang).Append(",");
                    sb.Append(tanggia.TyLeTang).Append(",");
                    sb.Append("'" + tanggia.Trangding_Date + "'");
                    sb.Append("),");

                }
                if (max != null)
                {
                    //Symbol,AvePrice,TyLeTang,KLGD,Trangding_Date
                    sb.Append("(");
                    sb.Append("'" + max.Symbol + "'").Append(",");
                    sb.Append(max.HighPrice).Append(",");
                    sb.Append(max.LowPrice).Append(",");
                    sb.Append(max.TyLeChenhLech).Append(",");
                    sb.Append("'" + max.Trangding_Date + "'");
                    sb.Append("),");

                }
                if (tp != null)
                {
                    //STT,Symbol,KyHanNam,GiaGDDong,LaiSuat,LoiSuat,KLGD,GTGD,Trangding_Date
                    sb.Append("(");
                    sb.Append(tp.STT).Append(",");
                    sb.Append("'" + tp.Symbol + "'").Append(",");
                    sb.Append(tp.KyHanNam).Append(",");
                    sb.Append(tp.GiaGDDong).Append(",");
                    sb.Append(tp.LaiSuat).Append(",");
                    sb.Append(tp.LoiSuat).Append(",");
                    sb.Append(tp.KLGD).Append(",");
                    sb.Append(tp.GTGD).Append(",");
                    sb.Append("'" + tp.Trangding_Date + "'");
                    sb.Append("),");

                }
                if (tc != null)
                {
                    //Symbol,KLMua_KL,KLBan_KL,KL_ChenhLech,GTMua_KL,GTBan_KL,
                    //KLMua_TT,KLBan_TT,KL_ChenhLech_TT,GTMua_TT,GTBan_TT,
                    //KLMua_TC,KLBan_TC,KL_ChenhLech_TC,GTMua_TC,GTBan_TC,Trangding_Date
                    sb.Append("(");
                    sb.Append("'" + tc.Symbol + "'").Append(",");
                    sb.Append(tc.KLMua_KL).Append(",");
                    sb.Append(tc.KLBan_KL).Append(",");
                    sb.Append(tc.KL_ChenhLech).Append(",");
                    sb.Append(tc.GTMua_KL).Append(",");
                    sb.Append(tc.GTBan_KL).Append(",");

                    sb.Append(tc.KLMua_TT).Append(",");
                    sb.Append(tc.KLBan_TT).Append(",");
                    sb.Append(tc.KL_ChenhLech_TT).Append(",");
                    sb.Append(tc.GTMua_TT).Append(",");
                    sb.Append(tc.GTBan_TT).Append(",");
                    sb.Append(tc.KLMua_TC).Append(",");
                    sb.Append(tc.KLBan_TC).Append(",");
                    sb.Append(tc.KL_ChenhLech_TC).Append(",");
                    sb.Append(tc.GTMua_TC).Append(",");
                    sb.Append(tc.GTBan_TC).Append(",");
                    sb.Append("'" + tc.Trangding_Date + "'");
                    sb.Append("),");

                }


                // ko exec sp, chi lay script de run bulk update sau nay 
                /* if (getScriptOnly)
                 {*/
                return new EDalResult() { Code = EDalResult.__CODE_SUCCESS, Message = EDalResult.__STRING_GET_SCRIPT, Data = sb.ToString() };
                // }

                /*  // 2. main			
                  result = await ExecuteSpNoQueryAsync(spName, dynamicParameters);


                  // return (neu sp ko tra error code,msg thi tu gan default)
                  return new EDalResult() { Code = EDalResult.__CODE_SUCCESS, Message = EDalResult.__STRING_SUCCESS, Data = result.Data };*/
            }
            catch (Exception ex)
            {
                // log error + buffer data
                //  this._cS6GApp.ErrorLogger.LogErrorContext(ex, ec);
                // error => return null
                return new EDalResult() { Code = -9997, Message = ex.Message, Data = null };
            }
        }
        public EDalResult Update_TTICBVCTDKGD_HNX(THONGTINCB_HNX ttcb_hnx, GIAODICHNHADAUTUNN_HNX gdndtnn_hnx, TKCUNGCAUTTCP_HNX cc_hnx,
                 Price_GDNKT_HNX price_hnx, KQGIAODICHCP_HNX kq_hnx, Chi_Tieu_HNX ct, Top10_CPGDMAX_HNX cpgdmax,
                 Top10_CPNYGTMAX_HNX cpnygtmax, Top10_CPMUAMAX_HNX cpmuamax, Top10_CPTANGPRICE_HNX cptangprice,
                 Top10_KLGDMAX_HNX klgdmax, Top10_CPGTVHMAX_HNX cpgtvhmax, Top10_CPBANMAX_HNX cpbanmax, Top10_CPGIAMPRICE_HNX cpgiamprice, KQGIAODICHCP_HNX2 kq_hnx2, bool getScriptOnly = false)
        {

            try
            {
                string spName = "";
                EDalResult result;
                //   DynamicParameters dynamicParameters = new DynamicParameters();
                StringBuilder sb = new StringBuilder();
                if (ttcb_hnx != null)
                {
                    sb.Append("(");
                    sb.Append(ttcb_hnx.STT).Append(",");
                    sb.Append("'" + ttcb_hnx.Symbol + "'").Append(",");
                    sb.Append(ttcb_hnx.PriceCloseAverage).Append(",");
                    sb.Append(ttcb_hnx.KLCPNY).Append(",");
                    sb.Append(ttcb_hnx.KLCPLH).Append(",");
                    sb.Append(ttcb_hnx.EPS).Append(",");
                    sb.Append(ttcb_hnx.EPS4).Append(",");
                    sb.Append(ttcb_hnx.PE).Append(",");
                    sb.Append(ttcb_hnx.ROE).Append(",");
                    sb.Append(ttcb_hnx.ROA).Append(",");
                    sb.Append(ttcb_hnx.GTTT).Append(",");
                    sb.Append("'" + ttcb_hnx.Trangding_Date + "'");
                    sb.Append("),");

                }
                if (gdndtnn_hnx != null)
                {
                    sb.Append("(");
                    sb.Append(gdndtnn_hnx.STT).Append(",");
                    sb.Append("'" + gdndtnn_hnx.Symbol + "'").Append(",");
                    sb.Append(gdndtnn_hnx.KLMUA_KL).Append(",");
                    sb.Append(gdndtnn_hnx.GTMUA_KL).Append(",");
                    sb.Append(gdndtnn_hnx.KLBAN_KL).Append(",");
                    sb.Append(gdndtnn_hnx.GTBAN_KL).Append(",");
                    sb.Append(gdndtnn_hnx.KLMUA_TT).Append(",");

                    sb.Append(gdndtnn_hnx.GTMUA_TT).Append(",");

                    sb.Append(gdndtnn_hnx.KLBAN_TT).Append(",");
                    sb.Append(gdndtnn_hnx.GTBAN_TT).Append(",");
                    sb.Append(gdndtnn_hnx.KLMUA_TC).Append(",");
                    sb.Append(gdndtnn_hnx.GTMUA_TC).Append(",");
                    sb.Append(gdndtnn_hnx.KLBAN_TC).Append(",");
                    sb.Append(gdndtnn_hnx.GTBAN_TC).Append(",");
                    sb.Append(gdndtnn_hnx.KLCK_MAX).Append(",");
                    sb.Append(gdndtnn_hnx.KLCK_NDTNN).Append(",");
                    sb.Append(gdndtnn_hnx.KLCK_CDPNG).Append(",");
                    sb.Append("'" + gdndtnn_hnx.Trangding_Date + "'");
                    sb.Append("),");

                }
                if (cc_hnx != null)
                {
                    sb.Append("(");
                    sb.Append(cc_hnx.STT).Append(",");
                    sb.Append("'" + cc_hnx.Symbol + "'").Append(",");
                    sb.Append(cc_hnx.SLDATMUA_KL).Append(",");
                    sb.Append(cc_hnx.KLDATMUA_KL).Append(",");
                    sb.Append(cc_hnx.SLDATBAN_KL).Append(",");
                    sb.Append(cc_hnx.KLDATBAN_KL).Append(",");
                    sb.Append(cc_hnx.SLDATMUA_TT).Append(",");
                    sb.Append(cc_hnx.KLDATMUA_TT).Append(",");
                    sb.Append(cc_hnx.SLDATBAN_TT).Append(",");
                    sb.Append(cc_hnx.KLDATBAN_TT).Append(",");
                    sb.Append(cc_hnx.SLDATMUA_TC).Append(",");
                    sb.Append(cc_hnx.KLDATMUA_TC).Append(",");
                    sb.Append(cc_hnx.SLDATBAN_TC).Append(",");
                    sb.Append(cc_hnx.KLDATBAN_TC).Append(",");
                    sb.Append(cc_hnx.KLDUMUA).Append(",");
                    sb.Append(cc_hnx.KLDUBAN).Append(",");
                    sb.Append(cc_hnx.KLTHUCHIEN).Append(",");
                    sb.Append(cc_hnx.GTTHUCHIEN).Append(",");
                    sb.Append("'" + cc_hnx.Trangding_Date + "'");
                    sb.Append("),");

                }
                if (kq_hnx != null)
                {
                    sb.Append("(");
                    sb.Append(kq_hnx.STT).Append(",");
                    sb.Append("'" + kq_hnx.Symbol + "'").Append(",");
                    sb.Append(kq_hnx.BasicPrice).Append(",");
                    sb.Append(kq_hnx.OpenPrice).Append(",");
                    sb.Append(kq_hnx.ClosePrice).Append(",");
                    sb.Append(kq_hnx.HighestPrice).Append(",");
                    sb.Append(kq_hnx.LowestPrice).Append(",");

                    sb.Append(kq_hnx.TDDiem).Append(",");

                    sb.Append(kq_hnx.TDPhanTram).Append(",");
                    sb.Append(kq_hnx.KLGDC_KL).Append(",");
                    sb.Append(kq_hnx.KLGDL_KL).Append(",");
                    sb.Append(kq_hnx.GTGDC_KL).Append(",");
                    sb.Append(kq_hnx.GTGDL_KL).Append(",");
                    sb.Append(kq_hnx.KLGDC_TT).Append(",");
                    sb.Append(kq_hnx.KLGDL_TT).Append(",");
                    sb.Append(kq_hnx.GTGDC_TT).Append(",");
                    sb.Append(kq_hnx.GTGDL_TT).Append(",");
                    sb.Append(kq_hnx.KLGD_TC).Append(",");
                    sb.Append(kq_hnx.TITRONG1).Append(",");
                    sb.Append(kq_hnx.GTGD_TC).Append(",");
                    sb.Append(kq_hnx.TITRONG2).Append(",");
                    sb.Append(kq_hnx.KLCPLH).Append(",");
                    sb.Append(kq_hnx.GTVHTT_GT).Append(",");
                    sb.Append(kq_hnx.GTVHTT_TT).Append(",");
                    sb.Append("'" + kq_hnx.Trangding_Date + "'");
                    sb.Append("),");

                }

                if (kq_hnx2 != null)
                {
                    sb.Append("(");
                    sb.Append(kq_hnx2.STT).Append(",");
                    sb.Append("'" + kq_hnx2.Symbol + "'").Append(",");
                    sb.Append(kq_hnx2.BasicPrice).Append(",");
                    sb.Append(kq_hnx2.OpenPrice).Append(",");
                    sb.Append(kq_hnx2.ClosePrice).Append(",");
                    sb.Append(kq_hnx2.HighestPrice).Append(",");
                    sb.Append(kq_hnx2.LowestPrice).Append(",");

                    sb.Append(kq_hnx2.TDDiem).Append(",");

                    sb.Append(kq_hnx2.TDPhanTram).Append(",");
                    sb.Append(kq_hnx2.KLGDC_KL).Append(",");
                    sb.Append(kq_hnx2.KLGDL_KL).Append(",");
                    sb.Append(kq_hnx2.GTGDC_KL).Append(",");
                    sb.Append(kq_hnx2.GTGDL_KL).Append(",");
                    sb.Append(kq_hnx2.KLGDC_TT).Append(",");
                    sb.Append(kq_hnx2.KLGDL_TT).Append(",");
                    sb.Append(kq_hnx2.GTGDC_TT).Append(",");
                    sb.Append(kq_hnx2.GTGDL_TT).Append(",");
                    sb.Append(kq_hnx2.KLGD_TC).Append(",");
                    sb.Append(kq_hnx2.TITRONG1).Append(",");
                    sb.Append(kq_hnx2.GTGD_TC).Append(",");
                    sb.Append(kq_hnx2.TITRONG2).Append(",");
                    sb.Append(kq_hnx2.KLCPLH).Append(",");
                    sb.Append(kq_hnx2.GTVHTT_GT).Append(",");
                    sb.Append(kq_hnx2.GTVHTT_TT).Append(",");
                    sb.Append(kq_hnx2.TrangThaiCK).Append(",");
                    sb.Append(kq_hnx2.TinhTrangCK).Append(",");
                    sb.Append(kq_hnx2.TrangThaiThucHienQuyen).Append(",");
                    sb.Append("'" + kq_hnx2.Trangding_Date + "'");
                    sb.Append("),");



                }

                if (ct != null)
                {
                    sb.Append("(");
                    sb.Append("'" + ct.Chi_Tieu + "'").Append(",");
                    sb.Append("'" + ct.Don_Vi + "'").Append(",");
                    sb.Append(ct.So_Lieu).Append(",");

                    sb.Append("'" + ct.Trangding_Date + "'");
                    sb.Append("),");


                }
                if (price_hnx != null)
                {

                    sb.Append("(");
                    sb.Append(price_hnx.STT).Append(",");
                    sb.Append("'" + price_hnx.Symbol + "'").Append(",");
                    sb.Append("'" + price_hnx.Market + "'").Append(",");
                    sb.Append(price_hnx.BasicPrice_HT).Append(",");
                    sb.Append(price_hnx.CeilingPrice_HT).Append(",");
                    sb.Append(price_hnx.FloorPrice_HT).Append(",");
                    sb.Append(price_hnx.BasicPrice_KT).Append(",");
                    sb.Append(price_hnx.CeilingPrice_KT).Append(",");
                    sb.Append(price_hnx.FloorPrice_KT).Append(",");

                    sb.Append("'" + price_hnx.Trangding_Date + "'");
                    sb.Append("),");

                }
                if (cpnygtmax != null)
                {
                    sb.Append("(");

                    sb.Append("'" + cpnygtmax.Symbol + "'").Append(",");

                    sb.Append(cpnygtmax.ClosePrice).Append(",");
                    sb.Append(cpnygtmax.KLGD).Append(",");
                    sb.Append(cpnygtmax.GTNY).Append(",");

                    sb.Append("'" + cpnygtmax.Trangding_Date + "'");
                    sb.Append("),");


                }

                if (cpgdmax != null)
                {
                    sb.Append("(");

                    sb.Append("'" + cpgdmax.Symbol + "'").Append(",");

                    sb.Append(cpgdmax.ClosePrice).Append(",");
                    sb.Append(cpgdmax.GTGD).Append(",");
                    sb.Append(cpgdmax.TyTrong).Append(",");

                    sb.Append("'" + cpgdmax.Trangding_Date + "'");
                    sb.Append("),");

                }


                if (cpmuamax != null)
                {
                    sb.Append("(");

                    sb.Append("'" + cpmuamax.Symbol + "'").Append(",");

                    sb.Append(cpmuamax.KLGD).Append(",");
                    sb.Append(cpmuamax.GTMUA).Append(",");
                    sb.Append(cpmuamax.KLNG).Append(",");

                    sb.Append("'" + cpmuamax.Trangding_Date + "'");
                    sb.Append("),");



                }
                if (cptangprice != null)
                {
                    sb.Append("(");

                    sb.Append("'" + cptangprice.Symbol + "'").Append(",");

                    sb.Append(cptangprice.ClosePrice).Append(",");
                    sb.Append(cptangprice.MucTang).Append(",");
                    sb.Append(cptangprice.TyLeTang).Append(",");
                    sb.Append(cptangprice.KLGD).Append(",");
                    sb.Append("'" + cptangprice.Trangding_Date + "'");
                    sb.Append("),");

                    //Symbol,ClosePrice,MucTang,TyLeTang,KLGD,Trangding_Date

                }
                // Top10_CPBANMAX_HNX cpbanmax, Top10_CPGIAMPRICE_HNX cpgiamprice
                if (klgdmax != null)
                {
                    sb.Append("(");

                    sb.Append("'" + klgdmax.Symbol + "'").Append(",");

                    sb.Append(klgdmax.ClosePrice).Append(",");
                    sb.Append(klgdmax.KLGD).Append(",");
                    sb.Append(klgdmax.TyTrong).Append(",");

                    sb.Append("'" + klgdmax.Trangding_Date + "'");
                    sb.Append("),");


                }
                if (cpgtvhmax != null)
                {
                    sb.Append("(");

                    sb.Append("'" + cpgtvhmax.Symbol + "'").Append(",");

                    sb.Append(cpgtvhmax.ClosePrice).Append(",");
                    sb.Append(cpgtvhmax.KLGD).Append(",");
                    sb.Append(cpgtvhmax.GTVHTT).Append(",");

                    sb.Append("'" + cpgtvhmax.Trangding_Date + "'");
                    sb.Append("),");

                }
                if (cpbanmax != null)
                {
                    sb.Append("(");

                    sb.Append("'" + cpbanmax.Symbol + "'").Append(",");

                    sb.Append(cpbanmax.KLBAN).Append(",");
                    sb.Append(cpbanmax.GTBAN).Append(",");
                    sb.Append(cpbanmax.KLNG).Append(",");

                    sb.Append("'" + cpbanmax.Trangding_Date + "'");
                    sb.Append("),");


                }
                if (cpgiamprice != null)
                {
                    sb.Append("(");

                    sb.Append("'" + cpgiamprice.Symbol + "'").Append(",");

                    sb.Append(cpgiamprice.ClosePrice).Append(",");
                    sb.Append(cpgiamprice.MucGiam).Append(",");
                    sb.Append(cpgiamprice.TyLeGiam).Append(",");
                    sb.Append(cpgiamprice.KLGD).Append(",");
                    sb.Append("'" + cpgiamprice.Trangding_Date + "'");
                    sb.Append("),");



                }
                // ko exec sp, chi lay script de run bulk update sau nay 
                /* if (getScriptOnly)
                 {*/
                return new EDalResult() { Code = EDalResult.__CODE_SUCCESS, Message = EDalResult.__STRING_GET_SCRIPT, Data = sb.ToString() };
                // }

                /*  // 2. main			
                  result = await ExecuteSpNoQueryAsync(spName, dynamicParameters);


                  // return (neu sp ko tra error code,msg thi tu gan default)
                  return new EDalResult() { Code = EDalResult.__CODE_SUCCESS, Message = EDalResult.__STRING_SUCCESS, Data = result.Data };*/
            }
            catch (Exception ex)
            {
                // log error + buffer data
                //  this._cS6GApp.ErrorLogger.LogErrorContext(ex, ec);
                // error => return null
                return new EDalResult() { Code = -9997, Message = ex.Message, Data = null };
            }
        }
        public EDalResult Update_TTICBVCTDKGD_HNX_2013(KQGIAODICHCP_HNX_2013 kq, NY_TTCP_2013 ttcp, KQGIAODICHCP_HNX_2013_2 kq2, bool getScriptOnly = false)
        {

            try
            {
                string spName = "";
                EDalResult result;
                //   DynamicParameters dynamicParameters = new DynamicParameters();
                StringBuilder sb = new StringBuilder();
                if (kq != null)
                {
                    //STT,Symbol,BasicPrice,OpenPrice,ClosePrice,HighestPrice,LowestPrice,TDDiem,TDPhanTram,
                    //KLGD_KL,GTGD_KL,KLGD_TT,GTGD_TT
                    //,KLGD_LL,GTGD_LL,KLGD_TC,TITRONG1,GTGD_TC,TITRONG2,KLCPLH,GTVHTT_GT,GTVHTT_TT,VonDL,Trangding_Date"
                    sb.Append("(");
                    sb.Append(kq.STT).Append(",");
                    sb.Append("'" + kq.Symbol + "'").Append(",");
                    sb.Append(kq.BasicPrice).Append(",");
                    sb.Append(kq.OpenPrice).Append(",");
                    sb.Append(kq.ClosePrice).Append(",");
                    sb.Append(kq.HighestPrice).Append(",");
                    sb.Append(kq.LowestPrice).Append(",");
                    sb.Append(kq.ClosePrice).Append(",");

                    sb.Append(kq.TDDiem).Append(",");

                    sb.Append(kq.TDPhanTram).Append(",");
                    sb.Append(kq.KLGD_KL).Append(",");
                    sb.Append(kq.GTGD_KL).Append(",");
                    sb.Append(kq.KLGD_TT).Append(",");
                    sb.Append(kq.GTGD_TT).Append(",");
                    sb.Append(kq.KLGD_LL).Append(",");
                    sb.Append(kq.GTGD_LL).Append(",");
                    sb.Append(kq.KLGD_TC).Append(",");
                    sb.Append(kq.TITRONG1).Append(",");
                    sb.Append(kq.GTGD_TC).Append(",");
                    sb.Append(kq.TITRONG2).Append(",");
                    sb.Append(kq.KLCPLH).Append(",");
                    sb.Append(kq.GTVHTT_GT).Append(",");
                    sb.Append(kq.GTVHTT_TT).Append(",");
                    sb.Append(kq.VonDL).Append(",");

                    sb.Append("'" + kq.Trangding_Date + "'");
                    sb.Append("),");


                }
                if (kq2 != null)
                {
                    //STT,Symbol,BasicPrice,OpenPrice,ClosePrice,HighestPrice,LowestPrice,TDDiem,TDPhanTram,
                    //KLGD_KL,GTGD_KL,KLGD_TT,GTGD_TT
                    //,KLGD_LL,GTGD_LL,KLGD_TC,TITRONG1,GTGD_TC,TITRONG2,KLCPLH,GTVHTT_GT,GTVHTT_TT,VonDL,Trangding_Date"
                    sb.Append("(");
                    sb.Append(kq2.STT).Append(",");
                    sb.Append("'" + kq2.Symbol + "'").Append(",");
                    sb.Append(kq2.BasicPrice).Append(",");
                    sb.Append(kq2.OpenPrice).Append(",");
                    sb.Append(kq2.ClosePrice).Append(",");
                    sb.Append(kq2.HighestPrice).Append(",");
                    sb.Append(kq2.LowestPrice).Append(",");
                    sb.Append(kq2.ClosePrice).Append(",");
                    sb.Append(kq2.GiaCoSo).Append(",");
                    sb.Append(kq2.TDDiem).Append(",");
                    sb.Append(kq2.TDPhanTram).Append(",");
                    sb.Append(kq2.KLGD_KL).Append(",");
                    sb.Append(kq2.GTGD_KL).Append(",");
                    sb.Append(kq2.KLGD_TT).Append(",");
                    sb.Append(kq2.GTGD_TT).Append(",");
                    sb.Append(kq2.KLGD_TC).Append(",");
                    sb.Append(kq2.TITRONG1).Append(",");
                    sb.Append(kq2.GTGD_TC).Append(",");
                    sb.Append(kq2.TITRONG2).Append(",");
                    sb.Append(kq2.KLCPLH).Append(",");
                    sb.Append(kq2.GTVHTT_GT).Append(",");
                    sb.Append(kq2.GTVHTT_TT).Append(",");
                    sb.Append(kq2.VonDL).Append(",");
                    sb.Append("'" + kq2.Trangding_Date + "'");
                    sb.Append("),");


                }
                if (ttcp != null)
                {
                    sb.Append("(");
                    sb.Append(ttcp.STT).Append(",");
                    sb.Append("'" + ttcp.Symbol + "'").Append(",");
                    sb.Append(ttcp.KLCP_NY).Append(",");
                    sb.Append(ttcp.KLCP_LH).Append(",");
                    sb.Append(ttcp.Co_Tuc_2013).Append(",");
                    sb.Append(ttcp.Co_Tuc_2014).Append(",");
                    sb.Append(ttcp.PE).Append(",");
                    sb.Append(ttcp.EPS2014).Append(",");
                    sb.Append(ttcp.ROE2014).Append(",");
                    sb.Append(ttcp.ROA2014).Append(",");
                    sb.Append(ttcp.BasicPrice_KT).Append(",");
                    sb.Append(ttcp.CeilingPrice_KT).Append(",");
                    sb.Append(ttcp.FloorPrice_KT).Append(",");
                    sb.Append("'" + ttcp.Trangding_Date + "'");
                    sb.Append("),");

                }
                // ko exec sp, chi lay script de run bulk update sau nay 
                /* if (getScriptOnly)
                 {*/
                return new EDalResult() { Code = EDalResult.__CODE_SUCCESS, Message = EDalResult.__STRING_GET_SCRIPT, Data = sb.ToString() };
                // }

                /*  // 2. main			
                  result = await ExecuteSpNoQueryAsync(spName, dynamicParameters);


                  // return (neu sp ko tra error code,msg thi tu gan default)
                  return new EDalResult() { Code = EDalResult.__CODE_SUCCESS, Message = EDalResult.__STRING_SUCCESS, Data = result.Data };*/
            }
            catch (Exception ex)
            {
                // log error + buffer data
                //  this._cS6GApp.ErrorLogger.LogErrorContext(ex, ec);
                // error => return null
                return new EDalResult() { Code = -9997, Message = ex.Message, Data = null };
            }
        }



        public EDalResult Update_TTICBVCTDKGD(THONGTINCB ttcb, GIAODICHNHADAUTUNN gdndtnn, TKCUNGCAUTTCP cc, KQGIAODICHCP kq, Chi_Tieu_UPCOM ct, Top10_CPGDT cpgdt, Top10_CPTPRICE cptprice, Top10_KLGDM klgdm, Top10_CPGIAMGIA cpgiamgia, Price_GDNKT price, KQGIAODICHCP2 kq2, bool getScriptOnly = false)
        {

            try
            {
                string spName = "";
                EDalResult result;
                // DynamicParameters dynamicParameters = new DynamicParameters();
                StringBuilder sb = new StringBuilder();
                if (ttcb != null)
                {
                    sb.Append("(");
                    sb.Append(ttcb.STT).Append(",");
                    sb.Append("'" + ttcb.Symbol + "'").Append(",");
                    sb.Append(ttcb.PriceCloseAverage).Append(",");
                    sb.Append(ttcb.KLCPNY).Append(",");
                    sb.Append(ttcb.KLCPLH).Append(",");
                    sb.Append(ttcb.EPS).Append(",");
                    sb.Append(ttcb.ROE).Append(",");
                    sb.Append(ttcb.ROA).Append(",");
                    sb.Append(ttcb.GTTT).Append(",");
                    sb.Append("'" + ttcb.Trangding_Date + "'");
                    sb.Append("),");

                }
                if (gdndtnn != null)
                {

                    sb.Append("(");
                    sb.Append(gdndtnn.STT).Append(",");
                    sb.Append("'" + gdndtnn.Symbol + "'").Append(",");
                    sb.Append(gdndtnn.KLMUA_KL).Append(",");
                    sb.Append(gdndtnn.GTMUA_KL).Append(",");
                    sb.Append(gdndtnn.KLBAN_KL).Append(",");
                    sb.Append(gdndtnn.GTBAN_KL).Append(",");
                    sb.Append(gdndtnn.KLMUA_TT).Append(",");
                    sb.Append(gdndtnn.GTMUA_TT).Append(",");
                    sb.Append(gdndtnn.KLBAN_TT).Append(",");
                    sb.Append(gdndtnn.GTBAN_TT).Append(",");
                    sb.Append(gdndtnn.KLMUA_TC).Append(",");
                    sb.Append(gdndtnn.GTMUA_TC).Append(",");
                    sb.Append(gdndtnn.KLBAN_TC).Append(",");
                    sb.Append(gdndtnn.GTBAN_TC).Append(",");
                    sb.Append(gdndtnn.KLCK_MAX).Append(",");
                    sb.Append(gdndtnn.KLCK_NDTNN).Append(",");
                    sb.Append(gdndtnn.KLCK_CDPNG).Append(",");
                    sb.Append("'" + gdndtnn.Trangding_Date + "'");
                    sb.Append("),");
                }
                if (cc != null)
                {
                    sb.Append("(");
                    sb.Append(cc.STT).Append(",");
                    sb.Append("'" + cc.Symbol + "'").Append(",");
                    sb.Append(cc.SLDATMUA_KL).Append(",");
                    sb.Append(cc.KLDATMUA_KL).Append(",");
                    sb.Append(cc.SLDATBAN_KL).Append(",");
                    sb.Append(cc.KLDATBAN_KL).Append(",");
                    sb.Append(cc.SLDATMUA_TT).Append(",");
                    sb.Append(cc.KLDATMUA_TT).Append(",");
                    sb.Append(cc.SLDATBAN_TT).Append(",");
                    sb.Append(cc.KLDATBAN_TT).Append(",");
                    sb.Append(cc.SLDATMUA_TC).Append(",");
                    sb.Append(cc.KLDATMUA_TC).Append(",");
                    sb.Append(cc.SLDATBAN_TC).Append(",");
                    sb.Append(cc.KLDATBAN_TC).Append(",");
                    sb.Append(cc.KLDUMUA).Append(",");
                    sb.Append(cc.KLDUBAN).Append(",");
                    sb.Append(cc.KLTHUCHIEN).Append(",");
                    sb.Append(cc.GTTHUCHIEN).Append(",");
                    sb.Append("'" + cc.Trangding_Date + "'");
                    sb.Append("),");

                }
                if (kq != null)
                {
                    sb.Append("(");
                    sb.Append(kq.STT).Append(",");
                    sb.Append("'" + kq.Symbol + "'").Append(",");
                    sb.Append(kq.BasicPrice).Append(",");
                    sb.Append(kq.OpenPrice).Append(",");
                    sb.Append(kq.ClosePrice).Append(",");
                    sb.Append(kq.HighestPrice).Append(",");
                    sb.Append(kq.LowestPrice).Append(",");
                    sb.Append(kq.ClosePrice).Append(",");
                    sb.Append(kq.AveragePrice).Append(",");
                    sb.Append(kq.TDDiem).Append(",");
                    sb.Append(kq.TDPhanTram).Append(",");
                    sb.Append(kq.KLGDC_KL).Append(",");
                    sb.Append(kq.KLGDL_KL).Append(",");
                    sb.Append(kq.GTGDC_KL).Append(",");
                    sb.Append(kq.GTGDL_KL).Append(",");
                    sb.Append(kq.KLGDC_TT).Append(",");
                    sb.Append(kq.KLGDL_TT).Append(",");
                    sb.Append(kq.GTGDC_TT).Append(",");
                    sb.Append(kq.GTGDL_TT).Append(",");
                    sb.Append(kq.KLGD_TC).Append(",");
                    sb.Append(kq.TITRONG1).Append(",");
                    sb.Append(kq.GTGD_TC).Append(",");
                    sb.Append(kq.TITRONG2).Append(",");
                    sb.Append(kq.KLCPLH).Append(",");
                    sb.Append(kq.GTVHTT_GT).Append(",");
                    sb.Append(kq.GTVHTT_TT).Append(",");
                    sb.Append("'" + kq.Trangding_Date + "'");
                    sb.Append("),");


                }
                if (kq2 != null)
                {
                    sb.Append("(");
                    sb.Append(kq2.STT).Append(",");
                    sb.Append("'" + kq2.Symbol + "'").Append(",");
                    sb.Append(kq2.BasicPrice).Append(",");
                    sb.Append(kq2.OpenPrice).Append(",");
                    sb.Append(kq2.HighestPrice).Append(",");
                    sb.Append(kq2.LowestPrice).Append(",");
                    sb.Append(kq2.AveragePrice).Append(",");
                    sb.Append(kq2.TDDiem).Append(",");
                    sb.Append(kq2.TDPhanTram).Append(",");
                    sb.Append(kq2.KLGDC_KL).Append(",");
                    sb.Append(kq2.KLGDL_KL).Append(",");
                    sb.Append(kq2.GTGDC_KL).Append(",");
                    sb.Append(kq2.GTGDL_KL).Append(",");
                    sb.Append(kq2.KLGDC_TT).Append(",");
                    sb.Append(kq2.KLGDL_TT).Append(",");
                    sb.Append(kq2.GTGDC_TT).Append(",");
                    sb.Append(kq2.GTGDL_TT).Append(",");
                    sb.Append(kq2.KLGD_TC).Append(",");
                    sb.Append(kq2.TITRONG1).Append(",");
                    sb.Append(kq2.GTGD_TC).Append(",");
                    sb.Append(kq2.TITRONG2).Append(",");
                    sb.Append(kq2.KLCPLH).Append(",");
                    sb.Append(kq2.GTVHTT_GT).Append(",");
                    sb.Append(kq2.GTVHTT_TT).Append(",");
                    sb.Append(kq2.TrangThaiCK).Append(",");
                    sb.Append(kq2.TinhTrangCK).Append(",");
                    sb.Append(kq2.TrangThaiThucHienQuyen).Append(",");
                    sb.Append("'" + kq2.Trangding_Date + "'");
                    sb.Append("),");
                }
                if (ct != null)
                {
                    sb.Append("(");
                    sb.Append("'" + ct.Chi_Tieu + "'").Append(",");
                    sb.Append("'" + ct.Don_Vi + "'").Append(",");
                    sb.Append(ct.So_Lieu).Append(",");
                    sb.Append("'" + ct.Trangding_Date + "'");
                    sb.Append("),");

                }
                if (cpgdt != null)
                {
                    sb.Append("(");
                    sb.Append("'" + cpgdt.Symbol + "'").Append(",");
                    sb.Append(cpgdt.GTGD).Append(",");
                    sb.Append(cpgdt.TyTrong).Append(",");
                    sb.Append("'" + cpgdt.Trangding_Date + "'");
                    sb.Append("),");
                }

                if (cptprice != null)
                {
                    sb.Append("(");
                    sb.Append("'" + cptprice.Symbol + "'").Append(",");
                    sb.Append(cptprice.MucTang).Append(",");
                    sb.Append(cptprice.TyLeTang).Append(",");
                    sb.Append(cptprice.KLGD).Append(",");
                    sb.Append("'" + cptprice.Trangding_Date + "'");
                    sb.Append("),");


                }

                if (klgdm != null)
                {
                    sb.Append("(");
                    sb.Append("'" + klgdm.Symbol + "'").Append(",");

                    sb.Append(klgdm.KLGD).Append(",");
                    sb.Append(klgdm.TyTrong).Append(",");

                    sb.Append("'" + klgdm.Trangding_Date + "'");
                    sb.Append("),");
                }

                if (cpgiamgia != null)
                {
                    sb.Append("(");
                    sb.Append("'" + cpgiamgia.Symbol + "'").Append(",");
                    sb.Append(cpgiamgia.MucGIAM).Append(",");
                    sb.Append(cpgiamgia.TyLeGiam).Append(",");
                    sb.Append(cpgiamgia.KLGD).Append(",");
                    sb.Append("'" + cpgiamgia.Trangding_Date + "'");
                    sb.Append("),");


                }
                if (price != null)
                {
                    sb.Append("(");
                    sb.Append(price.STT).Append(",");
                    sb.Append("'" + price.Symbol + "'").Append(",");
                    sb.Append("'" + price.Market + "'").Append(",");
                    sb.Append(price.BasicPrice_HT).Append(",");
                    sb.Append(price.CeilingPrice_HT).Append(",");
                    sb.Append(price.FloorPrice_HT).Append(",");
                    sb.Append(price.BasicPrice_KT).Append(",");
                    sb.Append(price.CeilingPrice_KT).Append(",");
                    sb.Append(price.FloorPrice_KT).Append(",");
                    sb.Append("'" + price.Trangding_Date + "'");
                    sb.Append("),");

                }

                // ko exec sp, chi lay script de run bulk update sau nay 
                /* if (getScriptOnly)
                 {*/
                return new EDalResult() { Code = EDalResult.__CODE_SUCCESS, Message = EDalResult.__STRING_GET_SCRIPT, Data = sb.ToString() };
                // }

                /*  // 2. main			
                  result = await ExecuteSpNoQueryAsync(spName, dynamicParameters);


                  // return (neu sp ko tra error code,msg thi tu gan default)
                  return new EDalResult() { Code = EDalResult.__CODE_SUCCESS, Message = EDalResult.__STRING_SUCCESS, Data = result.Data };*/
            }
            catch (Exception ex)
            {
                // log error + buffer data
                //  this._cS6GApp.ErrorLogger.LogErrorContext(ex, ec);
                // error => return null
                return new EDalResult() { Code = -9997, Message = ex.Message, Data = null };
            }
        }
        public EDalResult ExecBulkScript(string mssqlScript)
        {
            EDalResult mssqlResult = null;
            try
            {
                // update vao MSSQL
                mssqlResult = this.ExecuteScript(mssqlScript);


                // return data
                return new EDalResult()
                {
                    Code = mssqlResult.Code,
                    Message = mssqlResult.Message + "; ",
                    Data = mssqlResult.Data
                };
            }
            catch (Exception ex)
            {
                // log error + buffer data

                // return null
                return new EDalResult() { Code = EDalResult.__CODE_ERROR, Message = ex.Message, Data = null };
            }
        }
        public EDalResult ExecuteScript(string script)
        {
            try
            {
                EDalResult result;
                int affectedRowCount = ExecuteAsync(script);
                result = new EDalResult() { Code = EDalResult.__CODE_SUCCESS, Message = EDalResult.__STRING_SUCCESS, Data = affectedRowCount };
                return result;
            }
            catch (Exception ex)
            {

                // error => return null
                return new EDalResult() { Code = -9997, Message = ex.Message, Data = null };
            }
        }
        public int ExecuteAsync(string sql)
        {

            int affectedRows = 0;
            try
            {
                using (SqlConnection connection = new SqlConnection(this._configs.ConnectionString))
                {
                    connection.Open();

                    // exec async
                    affectedRows = connection.Execute(sql);
                    // log after: khong lay duoc data return, output sau khi exec; caller phai tu log neu can


                }
            }
            catch (Exception ex)
            {
                // log error + buffer data

            }
            return affectedRows;
        }
        public EBulkScript GetScriptNY2017(NY_KQGD kqgd, NY_ThongKeCC cc_hnx, NY_GDDTNN gddtnn, NY_TTCP ttcp)
        {
            EDalResult mssqlResult = null;
            // EDalResult oracleResult = null;

            try
            {
                // update vao MSSQL
                if (kqgd != null)
                {
                    mssqlResult = this.Update_2017(kqgd, null, null, null, true);
                }
                // update vao MSSQL
                if (cc_hnx != null)
                {
                    mssqlResult = this.Update_2017(null, cc_hnx, null, null, true);
                }
                if (gddtnn != null)
                {
                    mssqlResult = this.Update_2017(null, null, gddtnn, null, true);
                }
                if (ttcp != null)
                {
                    mssqlResult = this.Update_2017(null, null, null, ttcp, true);
                }

                //   mssqlResult =
                EBulkScript eBulk = new EBulkScript()
                {
                    MssqlScript = mssqlResult.Data.ToString()
                    //  OracleScript = oracleResult.Data.ToString()
                };

                // return data
                return eBulk;
            }
            catch (Exception ex)
            {

                // return null
                return null;
            }
        }
        public EBulkScript GetScriptTTCBHNX_2013(KQGIAODICHCP_HNX_2013 kqgd, NY_TTCP_2013 ny, KQGIAODICHCP_HNX_2013_2 kqgd2)
        {
            EDalResult mssqlResult = null;
            // EDalResult oracleResult = null;

            try
            {
                // update vao MSSQL
                if (kqgd != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX_2013(kqgd, null, null, true);
                }
                // update vao MSSQL
                if (ny != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX_2013(null, ny, null, true);
                }
                // update vao MSSQL
                if (kqgd2 != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX_2013(null, null, kqgd2, true);
                }

                //   mssqlResult =
                EBulkScript eBulk = new EBulkScript()
                {
                    MssqlScript = mssqlResult.Data.ToString()
                    //  OracleScript = oracleResult.Data.ToString()
                };

                // return data
                return eBulk;
            }
            catch (Exception ex)
            {

                // return null
                return null;
            }
        }
        public EBulkScript GetScriptTTCBHNX(THONGTINCB_HNX ttcb_hnx, GIAODICHNHADAUTUNN_HNX gdndtnn_hnx, TKCUNGCAUTTCP_HNX cc_hnx,
                 Price_GDNKT_HNX price_hnx, KQGIAODICHCP_HNX kq_hnx, Chi_Tieu_HNX ct, Top10_CPGDMAX_HNX cpgdmax,
                 Top10_CPNYGTMAX_HNX cpnygtmax, Top10_CPMUAMAX_HNX cpmuamax, Top10_CPTANGPRICE_HNX cptangprice,
                 Top10_KLGDMAX_HNX klgdmax, Top10_CPGTVHMAX_HNX cpgtvhmax, Top10_CPBANMAX_HNX cpbanmax, Top10_CPGIAMPRICE_HNX cpgiamprice, KQGIAODICHCP_HNX2 kq_hnx2)
        {
            EDalResult mssqlResult = null;
            // EDalResult oracleResult = null;

            try
            {
                // update vao MSSQL
                if (ttcb_hnx != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX(ttcb_hnx, null, null, null, null, null, null, null, null, null, null, null, null, null, null, true);
                }
                if (gdndtnn_hnx != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX(null, gdndtnn_hnx, null, null, null, null, null, null, null, null, null, null, null, null, null, true);
                }

                if (cc_hnx != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX(null, null, cc_hnx, null, null, null, null, null, null, null, null, null, null, null, null, true);
                }
                if (price_hnx != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX(null, null, null, price_hnx, null, null, null, null, null, null, null, null, null, null, null, true);
                }
                if (kq_hnx != null)
                {
                    // mssqlResult = Update_TTICBVCTDKGD_HNX_test(kq_hnx);
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX(null, null, null, null, kq_hnx, null, null, null, null, null, null, null, null, null, null, true);
                }
                if (kq_hnx2 != null)
                {

                    mssqlResult = this.Update_TTICBVCTDKGD_HNX(null, null, null, null, null, null, null, null, null, null, null, null, null, null, kq_hnx2, true);
                }

                if (ct != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX(null, null, null, null, null, ct, null, null, null, null, null, null, null, null, null, true);
                }
                if (cpgdmax != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX(null, null, null, null, null, null, cpgdmax, null, null, null, null, null, null, null, null, true);
                }

                if (cpnygtmax != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX(null, null, null, null, null, null, null, cpnygtmax, null, null, null, null, null, null, null, true);
                }
                if (cpmuamax != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX(null, null, null, null, null, null, null, null, cpmuamax, null, null, null, null, null, null, true);
                }
                if (cptangprice != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX(null, null, null, null, null, null, null, null, null, cptangprice, null, null, null, null, null, true);
                }
                if (klgdmax != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX(null, null, null, null, null, null, null, null, null, null, klgdmax, null, null, null, null, true);
                }
                ///
                if (cpgtvhmax != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX(null, null, null, null, null, null, null, null, null, null, null, cpgtvhmax, null, null, null, true);
                }
                if (cpbanmax != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX(null, null, null, null, null, null, null, null, null, null, null, null, cpbanmax, null, null, true);
                }
                if (cpgiamprice != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX(null, null, null, null, null, null, null, null, null, null, null, null, null, cpgiamprice, null, true);
                }

                //   mssqlResult =
                EBulkScript eBulk = new EBulkScript()
                {
                    MssqlScript = mssqlResult.Data.ToString()
                    //  OracleScript = oracleResult.Data.ToString()
                };

                // return data
                return eBulk;
            }
            catch (Exception ex)
            {

                // return null
                return null;
            }
        }
        //2010
        public EBulkScript GetScriptTTCBHNX2010(Top10CK_TANGGIA2010 tg, Top10CP_CLGMAX max, GD_TRAIPHIEU tp, GDTP_NDTNN tc)
        {
            EDalResult mssqlResult = null;
            // EDalResult oracleResult = null;

            try
            {

                // update vao MSSQL
                if (tg != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX_2010(tg, null, null, null, true);
                }
                if (max != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX_2010(null, max, null, null, true);
                }
                if (tp != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX_2010(null, null, tp, null, true);
                }
                if (tc != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX_2010(null, null, null, tc, true);
                }


                //   mssqlResult =
                EBulkScript eBulk = new EBulkScript()
                {
                    MssqlScript = mssqlResult.Data.ToString()
                    //  OracleScript = oracleResult.Data.ToString()
                };

                // return data
                return eBulk;
            }
            catch (Exception ex)
            {

                // return null
                return null;
            }
        }
        //2011
        public EBulkScript GetScriptTTCBHNX2011(KQGIAODICHCP2011 dkgd_hnx, TinhHinhDatLenh2011 thdl, NDTNN2011 ndtnn, KQGDCHITIET2011 kqct, KQGDTH2011 kqgdth, Top10CK_GTGDL gtgdl
            , Top10CK_KLGDL klgdl, Top10CP_GTNYL ctnyl, Top10CK_TANGGIA tanggia, Top10CK_GIAMGIA giamgia, Chi_Tieu_2011 ct, Top10CK_NDTNN ndtnns, KLGD_TOP2011_MR mr1,
           GTGD_TOP2011_MR mr2, TangGiam_TOP2011_MR mr3, CKNTDNN_TOP2011_MR mr4)
        {
            EDalResult mssqlResult = null;
            // EDalResult oracleResult = null;

            try
            {

                // update vao MSSQL
                if (dkgd_hnx != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX_2011(dkgd_hnx, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, true);
                }
                if (thdl != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX_2011(null, thdl, null, null, null, null, null, null, null, null, null, null, null, null, null, null, true);
                }

                if (ndtnn != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX_2011(null, null, ndtnn, null, null, null, null, null, null, null, null, null, null, null, null, null, true);
                }

                if (kqct != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX_2011(null, null, null, kqct, null, null, null, null, null, null, null, null, null, null, null, null, true);
                }

                if (kqgdth != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX_2011(null, null, null, null, kqgdth, null, null, null, null, null, null, null, null, null, null, null, true);
                }

                if (gtgdl != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX_2011(null, null, null, null, null, gtgdl, null, null, null, null, null, null, null, null, null, null, true);
                }
                if (klgdl != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX_2011(null, null, null, null, null, null, klgdl, null, null, null, null, null, null, null, null, null, true);
                }

                if (ctnyl != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX_2011(null, null, null, null, null, null, null, ctnyl, null, null, null, null, null, null, null, null, true);
                }
                if (tanggia != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX_2011(null, null, null, null, null, null, null, null, tanggia, null, null, null, null, null, null, null, true);
                }

                if (giamgia != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX_2011(null, null, null, null, null, null, null, null, null, giamgia, null, null, null, null, null, null, true);
                }

                if (ct != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX_2011(null, null, null, null, null, null, null, null, null, null, ct, null, null, null, null, null, true);
                }
                if (ndtnns != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX_2011(null, null, null, null, null, null, null, null, null, null, null, ndtnns, null, null, null, null, true);
                }

                if (mr1 != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX_2011(null, null, null, null, null, null, null, null, null, null, null, null, mr1, null, null, null, true);
                }
                if (mr2 != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX_2011(null, null, null, null, null, null, null, null, null, null, null, null, null, mr2, null, null, true);
                }
                if (mr3 != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX_2011(null, null, null, null, null, null, null, null, null, null, null, null, null, null, mr3, null, true);
                }
                if (mr4 != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_HNX_2011(null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, mr4, true);
                }

                //   mssqlResult =
                EBulkScript eBulk = new EBulkScript()
                {
                    MssqlScript = mssqlResult.Data.ToString()
                    //  OracleScript = oracleResult.Data.ToString()
                };

                // return data
                return eBulk;
            }
            catch (Exception ex)
            {

                // return null
                return null;
            }
        }
        //2011- upcom
        public EBulkScript GetScriptTTCBUPCOM_2011(UPCoM_KQGD_Phien_2011 cc)
        {
            EDalResult mssqlResult = null;
            // EDalResult oracleResult = null;

            try
            {

                // update vao MSSQL
                if (cc != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_UPCOM_2011(cc, true);
                }


                //   mssqlResult =
                EBulkScript eBulk = new EBulkScript()
                {
                    MssqlScript = mssqlResult.Data.ToString()
                    //  OracleScript = oracleResult.Data.ToString()
                };

                // return data
                return eBulk;
            }
            catch (Exception ex)
            {

                // return null
                return null;
            }
        }

        //2013 -UPCOM
        public EBulkScript GetScriptTTCBUPCOM_2013(UPCoM_GDNDTNN_Phien up1, UPCoM_KQGD_Phien up2, UPCoM_TKCC up3,
            UPCoM_CPDKGD_Phien up4)
        {
            EDalResult mssqlResult = null;
            // EDalResult oracleResult = null;

            try
            {

                // update vao MSSQL
                if (up1 != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_UPCOM_2013(up1, null, null, null, true);
                }
                if (up2 != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_UPCOM_2013(null, up2, null, null, true);
                }
                if (up3 != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_UPCOM_2013(null, null, up3, null, true);
                }
                if (up4 != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD_UPCOM_2013(null, null, null, up4, true);
                }


                //   mssqlResult =
                EBulkScript eBulk = new EBulkScript()
                {
                    MssqlScript = mssqlResult.Data.ToString()
                    //  OracleScript = oracleResult.Data.ToString()
                };

                // return data
                return eBulk;
            }
            catch (Exception ex)
            {

                // return null
                return null;
            }
        }


        //2011
        public EDalResult Update_TTICBVCTDKGD_UPCOM_2011(UPCoM_KQGD_Phien_2011 cc, bool getScriptOnly = false)
        {

            try
            {
                string spName = "";
                EDalResult result;
                //   DynamicParameters dynamicParameters = new DynamicParameters();
                StringBuilder sb = new StringBuilder();

                if (cc != null)
                {
                    sb.Append("(");
                    sb.Append("'" + cc.Trangding_Date + "'").Append(",");
                    sb.Append("'" + cc.GiaoDich + "'").Append(",");
                    sb.Append("'" + cc.Symbol + "'").Append(",");
                    sb.Append(cc.BasicPrice).Append(",");
                    sb.Append(cc.CellingPrice).Append(",");
                    sb.Append(cc.FloorPrice).Append(",");
                    sb.Append(cc.HighestPrice).Append(",");
                    sb.Append(cc.LowestPrice).Append(",");
                    sb.Append(cc.OpenPrice).Append(",");
                    sb.Append(cc.ClosePrice).Append(",");
                    sb.Append(cc.Gia_BQ).Append(",");
                    sb.Append(cc.KLGD_KL).Append(",");
                    sb.Append(cc.GTGD_KL).Append(",");
                    sb.Append(cc.HighestPrice_TT).Append(",");
                    sb.Append(cc.LowestPrice_TT).Append(",");
                    sb.Append(cc.KLGD_TT).Append(",");
                    sb.Append(cc.GTGD_TT).Append(",");
                    sb.Append(cc.KLGD_TC).Append(",");
                    sb.Append(cc.GTGD_TC).Append(",");
                    sb.Append(cc.Muc_VHTT).Append(",");
                    sb.Append(cc.KL_DKGD).Append(",");
                    sb.Append(cc.KLCPLH).Append(",");
                    sb.Append(cc.KLMUA).Append(",");
                    sb.Append(cc.GTMUA).Append(",");
                    sb.Append(cc.KLBAN).Append(",");
                    sb.Append(cc.GTBAN).Append(",");
                    sb.Append(cc.TongKLDPNG).Append(",");
                    sb.Append(cc.KLCDPNG);

                    sb.Append("),");

                }


                // ko exec sp, chi lay script de run bulk update sau nay 
                /* if (getScriptOnly)
                 {*/
                return new EDalResult() { Code = EDalResult.__CODE_SUCCESS, Message = EDalResult.__STRING_GET_SCRIPT, Data = sb.ToString() };
                // }

                /*  // 2. main			
                  result = await ExecuteSpNoQueryAsync(spName, dynamicParameters);


                  // return (neu sp ko tra error code,msg thi tu gan default)
                  return new EDalResult() { Code = EDalResult.__CODE_SUCCESS, Message = EDalResult.__STRING_SUCCESS, Data = result.Data };*/
            }
            catch (Exception ex)
            {
                // log error + buffer data
                //  this._cS6GApp.ErrorLogger.LogErrorContext(ex, ec);
                // error => return null
                return new EDalResult() { Code = -9997, Message = ex.Message, Data = null };
            }
        }
        //================//
        public EDalResult Update_TTICBVCTDKGD_UPCOM_2013(UPCoM_GDNDTNN_Phien up1, UPCoM_KQGD_Phien up2, UPCoM_TKCC up3,
            UPCoM_CPDKGD_Phien up4, bool getScriptOnly = false)
        {

            try
            {
                string spName = "";
                EDalResult result;
                //   DynamicParameters dynamicParameters = new DynamicParameters();
                StringBuilder sb = new StringBuilder();

                if (up3 != null)
                {
                    //STT,Symbol,SLDATMUA,KLDATMUA,SLDATBAN,KLDATBAN,CLMUABAN,Trangding_Date
                    sb.Append("(");
                    sb.Append(up3.STT).Append(",");
                    sb.Append("'" + up3.Symbol + "'").Append(",");
                    sb.Append(up3.SLDATMUA).Append(",");
                    sb.Append(up3.KLDATMUA).Append(",");
                    sb.Append(up3.SLDATBAN).Append(",");
                    sb.Append(up3.KLDATBAN).Append(",");
                    sb.Append(up3.CLMUABAN).Append(",");
                    sb.Append("'" + up3.Trangding_Date + "'");
                    sb.Append("),");

                }
                if (up4 != null)
                {
                    //STT,Symbol,KLCP_NY,KLCP_LH,Co_Tuc_2010,PE,EPS2010,ROE,ROA,BasicPrice_KT,CeilingPrice_KT,FloorPrice_KT,Trangding_Date
                    sb.Append("(");
                    sb.Append(up4.STT).Append(",");
                    sb.Append("'" + up4.Symbol + "'").Append(",");
                    sb.Append(up4.KLCP_NY).Append(",");
                    sb.Append(up4.KLCP_LH).Append(",");
                    sb.Append(up4.Co_Tuc_2010).Append(",");
                    sb.Append(up4.PE).Append(",");
                    sb.Append(up4.EPS2010).Append(",");
                    sb.Append(up4.ROE).Append(",");
                    sb.Append(up4.ROA).Append(",");
                    sb.Append(up4.BasicPrice_KT).Append(",");
                    sb.Append(up4.CeilingPrice_KT).Append(",");
                    sb.Append(up4.FloorPrice_KT).Append(",");
                    sb.Append("'" + up4.Trangding_Date + "'");
                    sb.Append("),");

                }
                if (up2 != null)
                {
                    //STT,Symbol,BasicPrice,HighestPrice,LowestPrice,OpenPrice,ClosePrice,AveragePrice,TDDiem,TDPhanTram,KLGD_KL,
                    //GTGD_KL,KLGD_TT,GTGD_TT,KLGD_TC,TITRONG1,GTGD_TC,TITRONG2,KLCPLH,GTVHTT_GT,GTVHTT_TT,Trangding_Date
                    sb.Append("(");
                    sb.Append(up2.STT).Append(",");
                    sb.Append("'" + up2.Symbol + "'").Append(",");
                    sb.Append(up2.BasicPrice).Append(",");
                    sb.Append(up2.HighestPrice).Append(",");
                    sb.Append(up2.LowestPrice).Append(",");
                    sb.Append(up2.OpenPrice).Append(",");
                    sb.Append(up2.ClosePrice).Append(",");
                    sb.Append(up2.AveragePrice).Append(",");
                    sb.Append(up2.TDDiem).Append(",");
                    sb.Append(up2.TDPhanTram).Append(",");
                    sb.Append(up2.KLGD_KL).Append(",");
                    sb.Append(up2.GTGD_KL).Append(",");
                    sb.Append(up2.KLGD_TT).Append(",");
                    sb.Append(up2.GTGD_TT).Append(",");
                    sb.Append(up2.KLGD_TC).Append(",");
                    sb.Append(up2.TITRONG1).Append(",");
                    sb.Append(up2.GTGD_TC).Append(",");
                    sb.Append(up2.TITRONG2).Append(",");
                    sb.Append(up2.KLCPLH).Append(",");
                    sb.Append(up2.GTVHTT_GT).Append(",");
                    sb.Append(up2.GTVHTT_TT).Append(",");
                    sb.Append("'" + up2.Trangding_Date + "'");
                    sb.Append("),");

                }

                if (up1 != null)
                {
                    //STT,Symbol,KLMUA_KL,GTMUA_KL,KLBAN_KL,GTBAN_KL,KLMUA_TT,GTMUA_TT,KLBAN_TT,GTBAN_TT,KLMUA_TC,GTMUA_TC,
                    //KLBAN_TC,GTBAN_TC,KLCK_MAX,KLCK_NDTNN,KLCK_CDPNG,Trangding_Date"
                    sb.Append("(");
                    sb.Append(up1.STT).Append(",");
                    sb.Append("'" + up1.Symbol + "'").Append(",");
                    sb.Append(up1.KLMUA_KL).Append(",");
                    sb.Append(up1.GTMUA_KL).Append(",");
                    sb.Append(up1.KLBAN_KL).Append(",");
                    sb.Append(up1.GTBAN_KL).Append(",");
                    sb.Append(up1.KLMUA_TT).Append(",");
                    sb.Append(up1.GTMUA_TT).Append(",");
                    sb.Append(up1.KLBAN_TT).Append(",");
                    sb.Append(up1.GTBAN_TT).Append(",");
                    sb.Append(up1.KLMUA_TC).Append(",");
                    sb.Append(up1.GTMUA_TC).Append(",");
                    sb.Append(up1.KLBAN_TC).Append(",");
                    sb.Append(up1.GTBAN_TC).Append(",");
                    sb.Append(up1.KLCK_MAX).Append(",");
                    sb.Append(up1.KLCK_NDTNN).Append(",");
                    sb.Append(up1.KLCK_CDPNG).Append(",");
                    sb.Append("'" + up1.Trangding_Date + "'");
                    sb.Append("),");

                }


                // ko exec sp, chi lay script de run bulk update sau nay 
                /* if (getScriptOnly)
                 {*/
                return new EDalResult() { Code = EDalResult.__CODE_SUCCESS, Message = EDalResult.__STRING_GET_SCRIPT, Data = sb.ToString() };
                // }

                /*  // 2. main			
                  result = await ExecuteSpNoQueryAsync(spName, dynamicParameters);


                  // return (neu sp ko tra error code,msg thi tu gan default)
                  return new EDalResult() { Code = EDalResult.__CODE_SUCCESS, Message = EDalResult.__STRING_SUCCESS, Data = result.Data };*/
            }
            catch (Exception ex)
            {
                // log error + buffer data
                //  this._cS6GApp.ErrorLogger.LogErrorContext(ex, ec);
                // error => return null
                return new EDalResult() { Code = -9997, Message = ex.Message, Data = null };
            }
        }
        public EBulkScript GetScriptTTCBUPCOM(THONGTINCB ttcb, GIAODICHNHADAUTUNN gdndtnn, TKCUNGCAUTTCP cc, KQGIAODICHCP kq, Chi_Tieu_UPCOM ct, Top10_CPGDT cpgdt, Top10_CPTPRICE cptprice, Top10_KLGDM klgdm, Top10_CPGIAMGIA cpgiamgia, Price_GDNKT price, KQGIAODICHCP2 kq2)
        {
            EDalResult mssqlResult = null;
            // EDalResult oracleResult = null;

            try
            {
                // update vao MSSQL
                if (ttcb != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD(ttcb, null, null, null, null, null, null, null, null, null, null, true);
                }
                if (gdndtnn != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD(null, gdndtnn, null, null, null, null, null, null, null, null, null, true);
                }

                if (cc != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD(null, null, cc, null, null, null, null, null, null, null, null, true);
                }
                if (kq != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD(null, null, null, kq, null, null, null, null, null, null, null, true);
                }
                if (ct != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD(null, null, null, null, ct, null, null, null, null, null, null, true);
                }
                if (cpgdt != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD(null, null, null, null, null, cpgdt, null, null, null, null, null, true);
                }
                if (cptprice != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD(null, null, cc, null, null, null, cptprice, null, null, null, null, true);
                }
                if (klgdm != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD(null, null, cc, null, null, null, null, klgdm, null, null, null, true);
                }
                if (cpgiamgia != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD(null, null, cc, null, null, null, null, null, cpgiamgia, null, null, true);
                }
                if (price != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD(null, null, cc, null, null, null, null, null, null, price, null, true);
                }
                if (kq2 != null)
                {
                    mssqlResult = this.Update_TTICBVCTDKGD(null, null, null, null, null, null, null, null, null, null, kq2, true);
                }

                //   mssqlResult =
                EBulkScript eBulk = new EBulkScript()
                {
                    MssqlScript = mssqlResult.Data.ToString()
                    //  OracleScript = oracleResult.Data.ToString()
                };

                // return data
                return eBulk;
            }
            catch (Exception ex)
            {

                // return null
                return null;
            }
        }

        public string GetScript(string sql, DynamicParameters parameters)
        {
            return $"{sql} {ParametersToString(parameters)}";
        }
        public string ParametersToString(DynamicParameters parameters)
        {
            var result = new StringBuilder();

            if (parameters != null)
            {
                var firstParam = true;
                var parametersLookup = (SqlMapper.IParameterLookup)parameters;
                foreach (var paramName in parameters.ParameterNames)
                {
                    if (!firstParam)
                    {
                        result.Append(", "); //"\r\n, "
                    }
                    firstParam = false;

                    result.Append('@');
                    result.Append(paramName);
                    result.Append(" = ");
                    try
                    {
                        var value = parametersLookup[paramName];// parameters.Get<dynamic>(paramName);

                        //((System.Collections.Generic.Dictionary<string, Dapper.DynamicParameters.ParamInfo>)parameters.parameters)["Name"].DbType

                        result.Append((value != null) ? $"'{value.ToString()}'" : "null");
                    }
                    catch
                    {
                        result.Append("unknown");
                    }
                }

            }
            return result.ToString();
        }

    }
}
