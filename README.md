# Aspiration-Point-Filtering-Project

## About this Project
This project is intended to filter scores for entrance exams of high schools With the input data imported from the excel file is enrollment information, test scores, aspirations of candidates and norms of high schools. 

The 9th graders are registered with 3 aspirations(type 1, 2, 3). Each high school will only be allowed to consider students with aspiration 2 when there are no more students with aspiration 1. Similarly, the superintendent will only consider students with  aspiration 3 if there are no more students with aspiration 2.

Enrollment quotas for each school are different and given in advance. Each student for admission will be based on scores from 4 subjects including Literature, Mathematics, History and Foreign Languages. In which the score of Literature and Math is multiplied by a factor of 2, the score of History is by the factor of 1, and the score of Foreign Language subject will be counted as a conditional score (Students who have a foreign language score below 2 will be considered paralyzed in that subject and cannot pass. to any school). In case the total score for admission is equal, the higher foreign language score will be prioritized for admission first.

## Data Description
### Input
The input data of this problem is an excel file in the path D:\PROJECT\input.xlsx, which includes some Sheets:
The first sheet named "thong_tin_xet_tuyen" includes a number of lines, each line includes column 1 is the student's registration number (in the 10th grade entrance exam), column 2 is the student's first and last name;

The second sheet named "diem_toan" also includes 2 columns, the first column is the registration number, the second column is the math test score. Similar are other sheets with the names “diem_van”, “diem_lich_su”, “diem_ngoai_ngu”. Note that the order of student registration numbers can be very different between sheets, students who miss any exam will receive a value of -100 and will not be considered for admission.

The fifth sheet includes information about applying to schools, there are 5 schools in Hanoi (Assume that, in reality there are over 40 high schools), this sheet includes 4 columns, the first column is the registration number of students, the second column is aspiration 1, the third column is aspiration 2, the fourth column is aspiration 3. The name of the high school for students to fill out will be coded into a number (in this project, the school name is numbered from 1 to 5).

The last sheet is "chi_tieu", This sheet consists of only 5 lines, in which the number of students from the respective schools are allowed to enroll in turn.
 
### Output
An Excel file will be automatically created with 6 sheets:

The first sheet is named "diem_chuan", will record the standard scores for high schools, each school is on 1 line, the first column is the school code (from 1 to 5), the second column is the aspiration 1 score, the following columns are points of aspirations 2, aspirations 3(if any).

The second sheet is “danh_sach_1”, in which the names of students who pass into school 1 are listed, sorted by scores from highest to lowest. The following sheets are similar to the matriculation lists of schools 2, 3, 4, 5.

## Contact: nguyenhuynhduc.work@gmail.com

