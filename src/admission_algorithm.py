class AdmissionAlgorithm:
    MAJOR_MAPPING = {
        'A': ['电子信息工程', '通信工程', '电磁场与无线技术'],
        'B': ['电子信息工程', '电磁场与无线技术', '通信工程'],
        'C': ['电磁场与无线技术', '电子信息工程', '通信工程'],
        'D': ['电磁场与无线技术', '通信工程', '电子信息工程'],
        'E': ['通信工程', '电子信息工程', '电磁场与无线技术'],
        'F': ['通信工程', '电磁场与无线技术', '电子信息工程']
    }
    
    def __init__(self, quotas):
        self.quotas = quotas.copy()
        self.remaining_quotas = quotas.copy()
    
    def process_admissions(self, student_data):
        """
        Process student admissions based on their rankings and preferences.
        
        Args:
            student_data (pd.DataFrame): DataFrame containing student information
                with columns: ['学号', '排名', '志愿选择']
        
        Returns:
            pd.DataFrame: DataFrame with admission results
        """
        # Sort students by ranking
        sorted_students = student_data.sort_values('排名')
        
        # Initialize results
        results = sorted_students.copy()
        results['录取专业'] = ''
        
        # Process each student in order of ranking
        for _, student in sorted_students.iterrows():
            assigned = False
            preferences = self.MAJOR_MAPPING[student['志愿选择']]
            
            # Try to assign student to their preferred major
            for major in preferences:
                if self.remaining_quotas[major] > 0:
                    self.remaining_quotas[major] -= 1
                    results.loc[student.name, '录取专业'] = major
                    assigned = True
                    break
            
            # If no major is available, mark as unassigned
            if not assigned:
                results.loc[student.name, '录取专业'] = '未分配'
        
        return results
    
    def get_remaining_quotas(self):
        """Return the remaining quotas for each major."""
        return self.remaining_quotas.copy()
    
    def reset_quotas(self):
        """Reset the remaining quotas to their original values."""
        self.remaining_quotas = self.quotas.copy() 