using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace DataObject
{
    public class data_object
    {
        public int id;
        public string station;
        public XYZ start_point;
        public double horizontal_angle, vertical_angle;
        public double offset, super_high_angle;

        public data_object()
        {
            horizontal_angle = 0.0;
            vertical_angle = 0.0;
        }

    }
    public class envelope_object
    {
        public List<XYZ> Dynamic_envelope = new List<XYZ>();
        public List<XYZ> Vehicle_envelope = new List<XYZ>();
        public List<XYZ> third_envelope = new List<XYZ>();
    }
    public class setting_station
    {
        public List<string[]> inverted_arc_station = new List<string[]>(); // 仰拱
        public List<string[]> track_bed_station = new List<string[]>(); // 道床
        public List<string[]> gutter_station = new List<string[]>(); // 排水溝
        public List<string[]> walk_way_station = new List<string[]>(); // 走道
        public List<string[]> third_rail_station = new List<string[]>(); // 第三軌
    }

    public class contact_channel_properties
    {
        public List<int> tunnel_type = new List<int>();
        public List<double> tunnel_elevation = new List<double>();
        public List<double> up_arc_radius = new List<double>();
        public List<double> up_arc_thickness = new List<double>();
        public List<double> dn_height = new List<double>();
        public List<double> dn_thickness = new List<double>();
        public List<double> hollow_width = new List<double>();
        public List<double> hollow_height = new List<double>();
        public List<double> hollow_depth = new List<double>();
        public List<double> hollow_bottom_thickness = new List<double>();
        public List<double> support_thickness = new List<double>();
        public List<double> tunnel_redius = new List<double>();


        //聯絡通道層數(由API寫入)
        public List<int> path_levels = new List<int>();
        public List<double> level_one = new List<double>();
        public List<double> level_two = new List<double>();
        public List<double> level_three = new List<double>();

        //防火門(由API寫入)
        public List<int> door_levels = new List<int>();
        public List<double> door_dis = new List<double>();
    }

    public class track_bed_properties
    {
        public double standard_top_height, standard_center_dis, standard_width, standard_depth, standard_elevation;
        public double float_top_height, float_elevation, float_width, float_depth;
        public double suppad_length, suppad_width, suppad_depth, suppad_side_dis;
        public double flat_top_height, flat_elevation, flat_width;
        //鋼軌&鋼軌基板
        public double rail_gauge, rail_face_width, rail_base_length, rail_base_width, rail_base_thickness, rail_base_slope, rail_base_dis;

        //第三軌
        public double third_steel_between_dis, third_track_elevation, third_bracket_spacing, bracket_steel_between_dis,
            third_bracket_length, third_bracket_width;
    }
    public class rebar_properties
    {
        public List<string> main_A_distance = new List<string>();
        public List<string> main_B_distance = new List<string>();
        public List<string> main_K_distance = new List<string>();
        public List<string> shear_A_distance = new List<string>();
        public List<string> shear_A_type = new List<string>();
        public List<string> shear_B_distance = new List<string>();
        public List<string> shear_B_type = new List<string>();
        public List<string> shear_K_distance = new List<string>();
        public List<string> shear_K_type = new List<string>();
        public double main_inner_protect;
        public double main_outer_protect;
        public double main_side_protect;
        public double A_shear_num, B_shear_num, K_shear_num;
    }
    public class properties_object
    {
        //A section
        public string id;
        public double inner_diameter;
        public double width;
        public double thickness;
        public int Type_A_q, Type_B_q, Type_K_q;
        public double displacement_angle;
        public int type_K_insert_type;
        public double Type_S_1;
        public double center_point;

        //B section
        /*Type_A-1(盾面)
        Type_A-2(盾面)
        Type_A-3(盾面)
        Type_B-1(盾面)
        Type_B-2(盾面)
        Type_K(盾面)
        Type_A-1(尾面)
        Type_A-2(尾面)
        Type_A-3(尾面)
        Type_B-1(尾面)
        Type_B-2(尾面)
        Type_K(尾面)
        接頭角度αr
        K環單邊內縮量*/

        public double Type_A_1, Type_A_2, Type_A_3;
        public double Type_B_1_head, Type_B_2_head;
        public double Type_B_1_tail, Type_B_2_tail;
        public double Type_K_1_head, Type_K_1_tail;
        public double displacement;
        public double U_steel_aisle_top, U_steel_aisle_bottom, U_steel_nonaisle_top, U_steel_nonaisle_bottom;


        //道床頂部高程
        public double standerd_top_height, float_top_height;

        //仰拱
        public string inverted_arc_slope;

        //U形排水溝(標準排水溝)
        public double u_cover_width, u_cover_length, u_cover_thick;
        public double u_groove_width, u_depth, u_radius;
        public double u_long_side_dis, u_short_side_dis;
        public double u_zn_steel_thick, u_zn_steel_long_pitch, u_zn_steel_short_pitch, u_zn_steel_pitch;

        //PVC管排水溝
        public double pvc_radius, pvc_thick;
        public double pvc_gutter_depth, pvc_gutter_radius, pvc_inverted_arc_radius, pvc_gutter_witdh;
        public double pvc_zn_cover_width, pvc_zn_cover_length, pvc_zn_cover_thick;
        public double pvc_zn_cover_steel_thick, pvc_zn_cover_long_pitch, pvc_zn_cover_short_pitch, pvc_zn_cover_pitch;

        //浮動式道床排水溝
        public double float_gutter_radius, float_gutter_depth;

        //走道
        public double walkway_top_elevation, walkway_edge_to_rail_center_dis;
        public double walkway_requirement_width, connection_thick;
        public string walkway_slope;
        public double walkway_protrusion_width, walkway_protrusion_depth, walkway_protrusion_bottom;

        //電纜溝槽
        public double cable_distance, cable_top_width, cable_bottom_width, cable_depth;

        //電纜溝槽混凝土蓋板
        public double cable_cover_width, cable_cover_thick, cable_cover_length;
        public double cable_cover_stick_out_dis,cable_cover_stick_out_thick, cable_cover_stick_out_width;

        //混凝土強度
        public double inverted_arc_concrete_strength, walk_way_concrete_strength;


        public double suppad_length, suppad_width, suppad_depth, suppad_side_dis;
        public double flat_top_height, flat_elevation, flat_width;
    }

    public class miles_data
    {
        public double sum_miles = 0;
        public string sum_miles_cal;

        public double inverted_floating = 0;
        public double inverted_standard = 0;
        public double inverted_flat = 0;
        public string inverted_floating_cal, inverted_standard_cal, inverted_flat_cal;

        public double track_bed_floating = 0;
        public double track_bed_standard = 0;
        public double track_bed_flat = 0;
        public string track_bed_floating_cal, track_bed_standard_cal, track_bed_flat_cal;

        public double drainage_floating = 0;
        public double drainage_standard = 0;
        public double drainage_flat = 0;
        public string drainage_floating_cal, drainage_standard_cal, drainage_flat_cal;

    }

}
