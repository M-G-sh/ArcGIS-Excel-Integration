import arcpy
import pandas as pd

# طلب مسار الملف من المستخدم
excel_file = input("أدخل مسار ملف Excel: ")
feature_class = input("أدخل مسار طبقة ArcGIS: ")

# قراءة البيانات من ملف Excel
try:
    df = pd.read_excel(excel_file)
    print("تم تحميل البيانات بنجاح:")
    print(df.head())  # عرض أول 5 صفوف للتحقق من البيانات
except Exception as e:
    print(f"خطأ في تحميل ملف Excel: {e}")
    exit()

# التحقق مما إذا كانت الطبقة موجودة
if arcpy.Exists(feature_class):
    print("الطبقة موجودة!")
else:
    print("الطبقة غير موجودة!")

# تأكد من أن الأعمدة المطلوبة فقط موجودة
# إزالة الأعمدة غير المطلوبة مثل 'OBJECTID' و 'Location'
columns_to_transfer = [col for col in df.columns if col not in ['OBJECTID', 'Location']]

# التحقق إذا كانت الحقول في Excel موجودة في الطبقة
fields_in_fc = [field.name for field in arcpy.ListFields(feature_class)]
missing_fields = [field for field in columns_to_transfer if field not in fields_in_fc]

# إذا كانت هناك حقول مفقودة في الطبقة، نقوم بإضافتها
if missing_fields:
    print(f"تمت إضافة الحقول التالية إلى الطبقة: {missing_fields}")
    for field in missing_fields:
        # تحديد نوع الحقل المناسب (يمكنك تعديله حسب البيانات)
        arcpy.AddField_management(feature_class, field, "DOUBLE")  # أو نوع آخر بناءً على البيانات

# قائمة لتخزين الأعمدة التي تحتوي على قيم مفقودة أو غير قابلة للتحويل
invalid_columns = []

# الآن، نبدأ بنقل البيانات من Excel إلى الطبقة
with arcpy.da.UpdateCursor(feature_class, columns_to_transfer) as cursor:
    for row, index in zip(cursor, df.index):
        for i, column in enumerate(columns_to_transfer):
            value = df[column].iloc[index]
            
            # التعامل مع القيم غير المتوافقة:
            if isinstance(value, str):  # إذا كانت القيمة نصية
                if value.strip().lower() == 'nan':  # إذا كانت القيمة "NaN" بنسق نصي
                    row[i] = None  # تعيين القيمة إلى None (فارغ)
                    if column not in invalid_columns:
                        invalid_columns.append(column)  # إضافة اسم العمود إلى القائمة
                else:
                    try:
                        # محاولة تحويل النص إلى رقم (للتأكد من إمكانية تحويله)
                        row[i] = float(value)
                    except ValueError:
                        # إذا لم يكن من الممكن التحويل، يتم تعيينه إلى None
                        row[i] = None
                        if column not in invalid_columns:
                            invalid_columns.append(column)  # إضافة اسم العمود إلى القائمة
            elif isinstance(value, (int, float)):  # إذا كانت القيمة عددية
                row[i] = value
            elif pd.isna(value):  # إذا كانت القيمة فارغة (NaN)
                row[i] = None  # تعيينها إلى None (فارغ)
                if column not in invalid_columns:
                    invalid_columns.append(column)  # إضافة اسم العمود إلى القائمة
            else:
                row[i] = None  # إذا كانت القيمة غير قابلة للتحويل، تعيينها إلى None
                if column not in invalid_columns:
                    invalid_columns.append(column)  # إضافة اسم العمود إلى القائمة

        cursor.updateRow(row)

# طباعة الأعمدة التي تحتوي على قيم مفقودة أو غير قابلة للتحويل
if invalid_columns:
    print(f"الأعمدة التي تحتوي على قيم مفقودة أو غير قابلة للتحويل: {invalid_columns}")
else:
    print("لا توجد قيم مفقودة أو غير قابلة للتحويل.")

print("تم نقل البيانات بنجاح إلى الطبقة!")
