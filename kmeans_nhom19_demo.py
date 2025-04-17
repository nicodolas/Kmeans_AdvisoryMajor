import numpy as np
import pandas as pd
from tkinter import Tk, Frame, Label, Canvas, Scrollbar

# Đọc file Excel
file_path = "Diem10SV.xlsx"  
df = pd.read_excel(file_path)

# Định nghĩa các môn học cho từng chuyên ngành
majors = {
    'Hệ thống thông tin': ['Cơ sở dữ liệu', 'Thực hành cơ sở dữ liệu', 'Hệ quản trị cơ sở dữ liệu', 
                            'Thực hành hệ quản trị cơ sở dữ liệu', 'Phân tích thiết kế hệ thống thông tin', 
                            'Thực hành phân tích thiết kế hệ thống thông tin'],
    'Công nghệ phần mềm': ['Ngôn ngữ lập trình', 
                            'Thực hành ngôn ngữ lập trình', 'Cấu trúc dữ liệu và giải thuật', 
                            'Thực hành cấu trúc dữ liệu và giải thuật','Lập trình hướng đối tượng', 'Thực hành lập trình hướng đối tượng'],
    'Mạng máy tính và truyền thông': ['Kiến trúc máy tính', 'Hệ điều hành', 'Mạng máy tính', 'Thực hành mạng máy tính','Quản trị mạng', 'Thực hành quản trị mạng'],
    'Thương mại điện tử': ['Thiết kế web', 'Thực hành thiết kế web', 'Cơ sở dữ liệu', 'Thực hành cơ sở dữ liệu','Đồ họa máy tính', 'Thực hành đồ họa máy tính','Thương mại điện tử ngành CNTT']
}

# Ngưỡng điểm phân lớp
CNPM_THRESHOLD=6.5
MMT_THRESHOLD=6
TMDT_THRESHOLD=5.5
HTTT_THRESHOLD=5
REMOVE_THRESHOLD = 4
#Ngưỡng phân cụm
UPPER_THRESHOLD = 7 
LOWER_THRESHOLD = 2

def filter_students(df, columns, threshold):
    """Lọc sinh viên có điểm >= threshold ở tất cả các môn của chuyên ngành."""
    valid_students = df[df[columns].min(axis=1) >= threshold]  # Sinh viên đủ điểm trong tất cả các môn
    invalid_students = df[df[columns].min(axis=1) < threshold]  # Sinh viên có môn dưới ngưỡng
    return valid_students, invalid_students

def kmeans_clustering(data, upper_threshold, lower_threshold):
    """
    Phân cụm bằng K-Means với xử lý outliers và 2 cụm: trên ngưỡng và dưới ngưỡng.
    """
    # Xử lý outliers: Clipping giá trị vượt ngưỡng
    data_clipped = data.clip(lower=lower_threshold, upper=upper_threshold, axis=1)

    # Khởi tạo centroid ban đầu (cụm trên ngưỡng và dưới ngưỡng)
    centroids = [
        [upper_threshold] * data_clipped.shape[1],  # Cụm trên ngưỡng
        [lower_threshold] * data_clipped.shape[1]   # Cụm dưới ngưỡng
    ]

    def euclidean_distance(point1, point2):
        """Tính khoảng cách Euclidean giữa 2 điểm."""
        return np.sqrt(np.sum((point1 - point2) ** 2))

    for _ in range(10):  # Số lần lặp tối đa
        # Gán mỗi điểm dữ liệu vào cụm gần nhất
        clusters = {0: [], 1: []}
        for idx, row in data_clipped.iterrows():
            distances = [
                euclidean_distance(row.values, centroid) for centroid in centroids
            ]
            cluster = np.argmin(distances)  # Chọn cụm có khoảng cách nhỏ nhất
            clusters[cluster].append(idx)

        # Cập nhật centroid dựa trên trung bình cụm
        new_centroids = []
        for cluster in clusters:
            if clusters[cluster]:
                # Tính trung bình các điểm trong cụm (nếu cụm không rỗng)
                new_centroids.append(data_clipped.loc[clusters[cluster]].mean(axis=0).values)
            else:
                # Giữ nguyên centroid nếu cụm rỗng
                new_centroids.append(centroids[cluster])

        # Kiểm tra hội tụ (nếu centroid không thay đổi đáng kể)
        if np.allclose(centroids, new_centroids, atol=1e-4):
            break

        centroids = new_centroids

    return clusters


def assign_major(df, majors, cluster_results):
    """
    Phân loại sinh viên vào chuyên ngành theo điểm trung bình, 
    ưu tiên từ cao đến thấp và dựa trên kết quả phân cụm KMeans.
    """

    assigned_students = []

    # Duyệt qua các chuyên ngành và các cụm sinh viên có khả năng học
    for major, clusters in cluster_results.items():
        valid_students = clusters['Có khả năng học']

        for idx, student in valid_students.iterrows():
            student_scores = {}

            # Tính điểm trung bình của sinh viên cho từng chuyên ngành
            for major, subjects in majors.items():
                avg_score = student[subjects].mean() if all(subject in student.index for subject in subjects) else 0
                student_scores[major] = avg_score

            # Kiểm tra số lần xuất hiện của từng điểm trung bình
            scores_count = pd.Series(list(student_scores.values())).value_counts()

            # Sắp xếp các chuyên ngành theo điểm trung bình (giảm dần)
            # và chỉ xét thứ tự ưu tiên nếu điểm trung bình bằng nhau
            sorted_majors = sorted(
                student_scores.items(),
                key=lambda x: (
                    -x[1],  # Điểm trung bình (giảm dần)
                    ['Công nghệ phần mềm', 'Mạng máy tính và truyền thông', 'Thương mại điện tử', 'Hệ thống thông tin'].index(x[0])  # Thứ tự ưu tiên
                ) if scores_count[x[1]] > 1 else (-x[1], 0)
            )
            print(sorted_majors)

            # Lặp qua các chuyên ngành theo thứ tự 
            for major, avg_score in sorted_majors:
                # Kiểm tra điều kiện cho từng chuyên ngành
                if major == 'Công nghệ phần mềm' and avg_score >= CNPM_THRESHOLD and all(student[subject] >= REMOVE_THRESHOLD for subject in majors[major]):
                    assigned_students.append((student['MSSV'], student['Họ Tên Sinh Viên'], major))
                    break
                elif major == 'Mạng máy tính và truyền thông' and avg_score >= MMT_THRESHOLD and all(student[subject] >= REMOVE_THRESHOLD for subject in majors[major]):
                    assigned_students.append((student['MSSV'], student['Họ Tên Sinh Viên'], major))
                    break
                elif major == 'Thương mại điện tử' and avg_score >= TMDT_THRESHOLD and all(student[subject] >= REMOVE_THRESHOLD for subject in majors[major]):
                    assigned_students.append((student['MSSV'], student['Họ Tên Sinh Viên'], major))
                    break
                elif major == 'Hệ thống thông tin' and avg_score >= HTTT_THRESHOLD and all(student[subject] >= REMOVE_THRESHOLD for subject in majors[major]):
                    assigned_students.append((student['MSSV'], student['Họ Tên Sinh Viên'], major))
                    break

    return assigned_students










# Hiển thị bảng trong cửa sổ cuộn
def create_scrollable_canvas(df, title, frame):
    """Tạo một canvas cuộn cho bảng"""
    canvas = Canvas(frame)
    scrollbar = Scrollbar(frame, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)
    table_frame = Frame(canvas)
    canvas.create_window((0, 0), window=table_frame, anchor="nw")
    scrollbar.grid(row=1, column=1, sticky="ns")  # Sử dụng grid thay vì pack
    canvas.grid(row=1, column=0, sticky="nsew")  # Sử dụng grid thay vì pack

    # Tạo bảng
    for i, row in df.iterrows():
        for j, value in enumerate(row):
            label = Label(table_frame, text=value, width=20, height=2, relief="solid", anchor="center")
            label.grid(row=i, column=j, padx=5, pady=5)

    table_frame.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))

def display_cluster_results():
    # Phân cụm sinh viên theo chuyên ngành
    results = {}
    for major, columns in majors.items():
        valid, invalid = filter_students(df, columns, REMOVE_THRESHOLD)
       
        if not valid.empty:
            clusters = kmeans_clustering(valid[columns], UPPER_THRESHOLD, LOWER_THRESHOLD)
            
            results[major] = {
                'Có khả năng học': valid.loc[clusters[0]],
                'Không có khả năng học': valid.loc[clusters[1]]
            }
        else:
            results[major] = {
                'Có khả năng học': pd.DataFrame(),
                'Không có khả năng học': pd.DataFrame()
            }
   
    # Khởi tạo giao diện Tkinter cho màn hình phân cụm
    root = Tk()
    root.title("Kết quả phân cụm sinh viên theo chuyên ngành")

    for i, (major, clusters) in enumerate(results.items()):
        frame = Frame(root)
        frame.grid(row=i // 2, column=i % 2, padx=10, pady=10, sticky="nsew")

        title_label = Label(frame, text=f"Chuyên ngành: {major}", font=("Arial", 14, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=5)

        if not clusters['Có khả năng học'].empty:
            create_scrollable_canvas(clusters['Có khả năng học'][['MSSV', 'Họ Tên Sinh Viên']], "Có khả năng học", frame)
        else:
            no_students_label = Label(frame, text="Không có sinh viên đủ điều kiện", font=("Arial", 12))
            no_students_label.grid(row=1, column=0, columnspan=2, pady=5)

    root.mainloop()

def display_assigned_students():
    # Phân lớp sinh viên vào chuyên ngành
    results = {}
    for major, columns in majors.items():
        valid, invalid = filter_students(df, columns, REMOVE_THRESHOLD)
       
        if not valid.empty:
            clusters = kmeans_clustering(valid[columns], UPPER_THRESHOLD, LOWER_THRESHOLD)
            results[major] = {
                'Có khả năng học': valid.loc[clusters[0]],
                'Không có khả năng học': valid.loc[clusters[1]]
            }
        else:
            results[major] = {
                'Có khả năng học': pd.DataFrame(),
                'Không có khả năng học': pd.DataFrame()
            }

    # Phân lớp sinh viên vào chuyên ngành
    assigned_students = assign_major(df, majors, results)

    # Chuyển kết quả thành DataFrame với MSSV, Họ Tên và Chuyên Ngành
    assigned_df = pd.DataFrame(assigned_students, columns=['MSSV', 'Họ Tên Sinh Viên', 'Chuyên Ngành'])

    # Loại bỏ các dòng trùng lặp
    assigned_df = assigned_df.drop_duplicates(subset=['MSSV'])

    # Kiểm tra nếu không có sinh viên nào được phân vào chuyên ngành
    if assigned_df.empty:
        root = Tk()
        root.title("Thông báo")

        label = Label(root, text="Bạn không phù hợp để học bất kì chuyên ngành nào.", font=("Arial", 14, "bold"), fg="red")
        label.pack(padx=20, pady=20)

        root.mainloop()
        return

    # Hiển thị bảng trong cửa sổ cuộn
    root = Tk()
    root.title("Kết quả phân lớp sinh viên vào chuyên ngành")

    # Tạo một hàm để tạo khung cho từng chuyên ngành
    def create_major_frame(major, students, row, col):
        frame = Frame(root)
        frame.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")

        title_label = Label(frame, text=f"Chuyên ngành: {major}", font=("Arial", 14, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=5)

        create_scrollable_canvas(students[['MSSV', 'Họ Tên Sinh Viên']], f"Chuyên ngành {major}", frame)

    # Điều chỉnh lưới 2x2
    for i, (major, students) in enumerate(assigned_df.groupby('Chuyên Ngành')):
        row = i // 2  # Tính số hàng
        col = i % 2   # Tính số cột
        create_major_frame(major, students, row, col)

    # Chạy giao diện Tkinter
    root.mainloop()




# Chạy màn hình phân cụm trước
display_cluster_results()

# Sau khi phân cụm hoàn tất, hiển thị màn hình phân lớp
display_assigned_students()
