public void ReadExcel()
        {
            Dictionary<int, DataTable> sprintSheets = new Dictionary<int, DataTable>();

            using (OleDbConnection conn = new OleDbConnection())
            {
                string path = "sprints.xlsx";

                string connStr = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Mode=Read;Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;MAXSCANROWS=0'", path);

                conn.ConnectionString = connStr;

                conn.Open();

                DataTable dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (dt == null)
                {
                    //return null;
                }

                List<string> excelSheets = new List<string>();

                // Add the sheet name to the string array.
                foreach (DataRow row in dt.Rows)
                {
                    excelSheets.Add(row["TABLE_NAME"].ToString());
                }

                // Loop through all of the sheets if you want too...
                for (int j = 0; j < 9; j++)
                {
                    // Query each excel sheet.

                    using (OleDbCommand comm = new OleDbCommand())
                    {
                        comm.CommandText = string.Format("Select * from [{0}]", excelSheets[j]);
                        comm.Connection = conn;
                        using (OleDbDataAdapter da = new OleDbDataAdapter())
                        {
                            DataTable sheetData = new DataTable();

                            da.SelectCommand = comm;
                            da.Fill(sheetData);
                            sprintSheets.Add(j, sheetData);
                        }
                    }
                }
            }

            foreach (KeyValuePair<int, DataTable> pair in sprintSheets)
            {
                SaveSprintData(pair.Key, pair.Value);
            }
        }

        public void SaveSprintData(int index, DataTable sheetData)
        {
            int sprintNo = index + 29;

            foreach (DataRow dr in sheetData.Rows)
            {
                if (Convert.ToString(dr[0]) == "Talep No" || Convert.ToString(dr[2]) == string.Empty || Convert.ToString(dr[0]).StartsWith("Sprint"))
                {
                    continue;
                }

                string talepNo = Convert.ToString(dr[0]);
                string btTaskno = Convert.ToString(dr[1]);
                string talepAdi = Convert.ToString(dr[2]);
                string analist = Convert.ToString(dr[3]);
                string[] analistler = analist.Split('\\');
                string yazilimci = Convert.ToString(dr[4]);
                string[] yazilimcilar = yazilimci.Split('\\');

                string state = Convert.ToString(dr[8]);
                string note = Convert.ToString(dr[9]);

                string taskType = talepNo.StartsWith("TLP") ? "Talep" : talepNo.StartsWith("INC") ? "INC" : "Diğer";

                foreach (string analizci in analistler)
                {
                    InsertToDb(sprintNo, talepNo, btTaskno, talepAdi, analizci, state, taskType, note);
                }
                foreach (string developer in yazilimcilar)
                {
                    InsertToDb(sprintNo, talepNo, btTaskno, talepAdi, developer, state, taskType, note);
                }
            }
        }

        public void InsertToDb(int sprintNo, string talepNo, string btTaskNo, string talepAdi, string assignee, string state, string taskType, string note)
        {
            string connStr = @"Server=(localdb)\MSSQLLocalDB;Database=Work;Integrated Security=true;";

            using (SqlConnection conn = new SqlConnection(connStr))
            using (SqlCommand comm = new SqlCommand())
            {
                comm.Connection = conn;
                comm.CommandType = CommandType.StoredProcedure;
                comm.CommandText = "dbo.ins_SprintAssignment";

                comm.Parameters.Add(new SqlParameter()
                {
                    ParameterName = "SprintNo",
                    SqlDbType = SqlDbType.Int,
                    Value = sprintNo
                });

                comm.Parameters.Add(new SqlParameter()
                {
                    ParameterName = "TalepNo",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 15,
                    Value = talepNo
                });

                comm.Parameters.Add(new SqlParameter()
                {
                    ParameterName = "BTTaskNo",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 15,
                    Value = btTaskNo
                });

                comm.Parameters.Add(new SqlParameter()
                {
                    ParameterName = "TalepAdi",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 150,
                    Value = talepAdi
                });

                comm.Parameters.Add(new SqlParameter()
                {
                    ParameterName = "Assignee",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Value = assignee
                });

                comm.Parameters.Add(new SqlParameter()
                {
                    ParameterName = "TaskStatu",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 20,
                    Value = state
                });

                comm.Parameters.Add(new SqlParameter()
                {
                    ParameterName = "TaskType",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 10,
                    Value = taskType
                });

                comm.Parameters.Add(new SqlParameter()
                {
                    ParameterName = "Note",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 500,
                    Value = note
                });

                conn.Open();

                comm.ExecuteNonQuery();
            }
        }
