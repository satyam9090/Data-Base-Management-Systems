# Data-Base-Management-Systems
Most colleges have a number of different courses and each course has a number of subjects. Now there are limited faculties, each faculty teaching more than one subjects. So now the time table needed to schedule the faculty at provided time slots in such a way that their timings do not overlap and the time table schedule makes best use of all faculty subject demands. The existing algorithm is long and the process becomes tedious. Our algorithm aims to provide a hassle-free way to generate a timetable for the faculties by using only one pre-existing table from the faculty wish list and details about the courses. The administrators need not worry about time clashes and there is no need for him to perform any permutations and combinations. An effective timetable is crucial for the satisfaction of enormous requirement and the efficient utilization of human and space resources, which make it an optimization problem. By the common hit and trial method, a solution is not guaranteed. We use a priority based allocation method, which allocates courses to faculties based on their wish list. Keeping in mind the maximum and minimum credits to be allocated, and also taking into consideration the maximum number of subjects that can be allocated to each faculty, we assign the minimum credits to all faculties. Then moving as per the priority list, based on seniority, we assign the remaining courses to the faculties. This method ensures that all faculties are given the minimum credits required, according to their preference. Also, no courses are repeated for a faculty. 
