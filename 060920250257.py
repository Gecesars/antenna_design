def _get_gain_cut(self, frequency: float, cut: str, fixed_angle_deg: float):
    """Get far-field gain data for a specific cut (theta or phi) at a given frequency.
    
    Args:
        frequency: Frequency in GHz
        cut: Either 'theta' or 'phi'
        fixed_angle_deg: Fixed angle for the cut in degrees
        
    Returns:
        Tuple of (angle_values, gain_values) or (None, None) if data is not available
    """
    try:
        # Create a unique report name based on the cut
        report_name = f"Gain_{cut}_cut_{fixed_angle_deg}deg"
        
        # Define the expression for gain in dB
        expression = "dB(GainTotal)"
        
        if cut == "theta":
            # Theta cut: vary theta, fix phi
            variations = {
                "Freq": f"{frequency}GHz",
                "Phi": f"{fixed_angle_deg}deg",
                "Theta": "All"
            }
            primary_sweep = "Theta"
        else:
            # Phi cut: vary phi, fix theta
            variations = {
                "Freq": f"{frequency}GHz",
                "Theta": f"{fixed_angle_deg}deg",
                "Phi": "All"
            }
            primary_sweep = "Phi"
        
        # Create the far field report
        report = self.hfss.post.reports_by_category.far_field(
            expressions=[expression],
            context="Infinite Sphere1",
            primary_sweep_variable=primary_sweep,
            primary_sweep_range="All",
            secondary_sweep_variable=None,
            report_category="Far Fields",
            setup_name="Setup1 : Sweep1",
            variations=variations,
            name=report_name
        )
        
        # Get the solution data
        solution_data = report.get_solution_data()
        
        if solution_data and hasattr(solution_data, 'primary_sweep_values'):
            angles = np.array(solution_data.primary_sweep_values)
            gains = np.array(solution_data.data_real())[0]  # Get first (and only) expression data
            
            # Ensure we have valid data
            if len(angles) > 0 and len(angles) == len(gains):
                return angles, gains
        
        return None, None
        
    except Exception as e:
        self.log_message(f"Error getting {cut} cut data: {e}")
        self.log_message(f"Traceback: {traceback.format_exc()}")
        return None, None