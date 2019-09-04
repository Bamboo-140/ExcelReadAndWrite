//传入DataGridView
        /// <summary>
        /// 输出数据到Excel
        /// </summary>
        /// <param name="dataGridView">DataGridView</param>
        /// <param name="includeHeader">是否包含字段名</param>
        /// <param name="savePath">保存路径</param>
        /// <returns>成功TRUE/失败FALSE</returns>
        public static bool ExportExcel(DataGridView dataGridView, bool includeHeader, string savePath)
        {
            bool succeed = true;
            try
            {
                //创建工作薄
                IWorkbook wbook = null;
                //创建工作表
                ISheet sheet = null;
                wbook = ICreateIWorkBook(savePath, null);
                sheet = wbook.CreateSheet("Sheet1");           //创建工作表                
                int columnCount = dataGridView.ColumnCount;    //获取行数
                int rowCount = dataGridView.RowCount;          //获取列数
                int rowIndex = 0;
                IRow row = null;
                ICell cell;
                int start = 0;    //是否包含表头行数
                //dataGridView是否有数据
                if (rowCount <= 0)
                {
                    succeed = false;
                    Exception error = new Exception("无可用于导出的数据");
                    throw error;
                }
                //是否包含表头
                if (includeHeader)
                {
                    row = sheet.CreateRow(rowIndex);
                    //输入标题表头行
                    for (int i = 0; i < columnCount; i++)
                    {
                        cell = row.CreateCell(i);
                        cell.SetCellValue(dataGridView.Columns[i].HeaderText);
                    }
                    ++start;     //如果包含表头行索引累加
                }

                //DataGridView有多少行
                for (; rowIndex < rowCount; ++rowIndex)
                {
                    row = sheet.CreateRow(rowIndex + start);    //创建行
                    for (int cellIndex = 0; cellIndex < columnCount; ++cellIndex)
                    {
                        Type type = dataGridView[cellIndex, rowIndex].ValueType;
                        object value = dataGridView[cellIndex, rowIndex].Value;
                        cell = row.CreateCell(cellIndex);     //创建单元格
                        //写入数据
                        if (type == typeof(Int32))
                        {
                            cell.SetCellValue((int)dataGridView[cellIndex, rowIndex].Value);
                        }
                        else if (type == typeof(Double))
                        {
                            cell.SetCellValue((double)value);
                        }
                        else if (type == typeof(float))
                        {
                            cell.SetCellValue((float)value);
                        }
                        else if (type == typeof(DateTime))
                        {
                            cell.SetCellValue((DateTime)value);
                        }
                        else if (type == typeof(Boolean))
                        {
                            cell.SetCellValue((bool)value);
                        }
                        else
                        {
                            cell.SetCellValue((string)value);
                        }
                    }
                }
                //保存的文件流保存或创建
                FileStream fs = new FileStream(savePath, FileMode.OpenOrCreate);
                wbook.Write(fs);   //写入文件
                wbook.Close();
                fs.Close();

            }
            catch (Exception ex)
            {
                succeed = false;
                throw ex;
            }
            return succeed;
        }


        /// <summary>
        /// 根据文件格式返回表文件
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="fs">文件流</param>
        /// <returns>IWorkBook 表</returns>
        private static IWorkbook ICreateIWorkBook(string filePath, FileStream fs)
        {
            IWorkbook wbook = null;
            //读取用户选择的保存格式
            string extesion = Path.GetExtension(filePath).ToLower();
            if (extesion.Equals(".xlsx"))
            //if( saveFile.FileName.IndexOf(".xlsx") > 0);
            {
                if (null != fs)
                {
                    wbook = new XSSFWorkbook(fs);
                }
                else
                {
                    wbook = new XSSFWorkbook();
                }
            }
            else if (extesion.Equals(".xls"))
            //else if (saveFile.FileName.IndexOf(".xls") > 0) ;    //这种方法也可以实现但.xlsx必须在前面
            {
                if (null != fs)
                {
                    wbook = new HSSFWorkbook(fs);
                }
                else
                {
                    wbook = new HSSFWorkbook();
                }
            }
            else    //如果不是上述两种格式,因为设置了过滤器几乎不会出现下面的情况
            {
                //MessageBox.Show("对不起，您所输入的文件格式不受支持！");
                //return null;
                Exception error = new Exception("文件格式不正确!");
                throw error;
            }
            return wbook;
        }


//如有问题欢迎大家指正
