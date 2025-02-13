  /** 
   * 試験の種類を列挙するEnum
   * @enum {number} 
   */
  const TestType = {
    NORMAL : "NORMAL",
    ABNORMAL : "ABNORMAL",
  };

  /** 
   * エラーコードの種類を列挙するEnum
   * @enum {String} 
   */
  const ErrorCodeType = {
    Min : "Min",
    Max : "Max",
    Size : "Size",
    NotNull : "NotNull",
    Past : "Past",
    TypeMismatch : "typeMismatch",
    Invalid: "Invalid",
  };

// 定義値の位置
const COLUMNS_Physics = 14
const COLUMNS_Type = 15
const COLUMNS_Required = 16
const COLUMNS_Min = 17
const COLUMNS_Max = 18
const COLUMNS_Enum = 19

// 項目チェック
const ANO_COLUMNS_logicalName = 0
const ANO_COLUMNS_Anotation = 1
const ANO_COLUMNS_ErrorMessage = 2

//
const FixedFiledNum = 2;
