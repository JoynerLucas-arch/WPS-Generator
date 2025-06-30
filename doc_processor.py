from template_analyzer import TemplateAnalyzer


class DocumentProcessor:
    @staticmethod
    def insert_data_to_no_content_point(p_d: dict):
        """校验模板并在预处理阶段插入无内容标签的信息"""
        p_t = p_d['type']
        if p_t in TemplateAnalyzer.insert_point_no_content_types:
            no_content_label = TemplateAnalyzer.registered_labels[p_t]
            no_content_label.insert_data_to_point(p_d, None, TemplateAnalyzer.static_datas)
            return True
        return False

    @staticmethod
    def solve_content_labels(insert_points, datas):
        """处理有内容类型插入点，包括表格中的内容，增强对表格内容的处理"""
        no_data_points = {}
        
        # 检查图片标签
        image_tags = []  # 存储所有图片标签信息
        for point_name, point_data in insert_points.items():
            if isinstance(point_data, list):
                for pd in point_data:
                    if pd['type'] == 'image':
                        image_tags.append((point_name, pd['text']))
            elif point_data['type'] == 'image':
                image_tags.append((point_name, point_data['text']))
                
        # 检查图片数据
        image_data = {}  # 存储所有图片数据
        for key, value in datas.items():
            if isinstance(value, tuple) and len(value) == 2 and isinstance(value[1], str):
                if value[1].endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                    image_data[key] = value
        
        # 如果有图片标签和图片数据，尝试智能匹配
        if image_tags and image_data:
            mapped_data = DocumentProcessor._smart_match_image_tags(image_tags, image_data, datas)
            
            # 合并原始数据和智能匹配的数据
            merged_datas = datas.copy()
            merged_datas.update(mapped_data)
        else:
            merged_datas = datas
        
        # 先创建标签到数据的映射，避免同一标签重复处理多次
        tag_data_map = {}
        
        # 验证数据有效性
        for point_name, point_data in insert_points.items():
            if point_name in merged_datas:
                data = merged_datas[point_name]
                
                # 验证数据类型
                if isinstance(point_data, list):
                    valid = False
                    for pd in point_data:
                        label = TemplateAnalyzer.registered_labels[pd['type']]
                        if label.check_data_type(data):
                            valid = True
                            break
                    
                    if not valid:
                        print(f"警告：标签 '{point_name}' 的数据类型不匹配，已跳过处理")
                        no_data_points[point_name] = point_data
                        continue
                else:
                    label = TemplateAnalyzer.registered_labels[point_data['type']]
                    if not label.check_data_type(data):
                        print(f"警告：标签 '{point_name}' 的数据类型不匹配，已跳过处理")
                        no_data_points[point_name] = point_data
                        continue
                
                tag_data_map[point_name] = data
            else:
                # 检查是否是大小写问题
                lower_keys = {k.lower(): k for k in merged_datas.keys()}
                if point_name.lower() in lower_keys:
                    actual_key = lower_keys[point_name.lower()]
                    print(f"标签名可能存在大小写问题: 模板中是 '{point_name}', 数据中是 '{actual_key}'")
                
                no_data_points[point_name] = point_data

        # 处理每个插入点
        for point_name, data in tag_data_map.items():
            try:
                point_data = insert_points[point_name]
                
                # 如果是同名多个标签的情况
                if isinstance(point_data, list):
                    # 检查是否为图片标签，若是，则确保不会重复插入图片到多个地方
                    is_image_type = False
                    table_image_points = []
                    non_table_image_points = []
                    
                    # 首先检查是否有图片类型的标签，并区分表格中的和非表格中的
                    for pd in point_data:
                        if pd['type'] == 'image':
                            is_image_type = True
                            if 'cell' in pd:
                                table_image_points.append(pd)
                            else:
                                non_table_image_points.append(pd)
                    
                    # 如果是图片类型，优先处理表格中的图片标签，其次处理非表格中的
                    if is_image_type:
                        # 处理所有表格中的图片标签
                        for pd in table_image_points:
                            label = TemplateAnalyzer.registered_labels[pd['type']]
                            try:
                                label.insert_data_to_point(pd, data, TemplateAnalyzer.static_datas)
                            except Exception as e:
                                print(f"错误：处理标签 {pd['text']} 时发生错误: {str(e)}")
                        
                        # 处理所有非表格中的图片标签
                        for pd in non_table_image_points:
                            label = TemplateAnalyzer.registered_labels[pd['type']]
                            try:
                                label.insert_data_to_point(pd, data, TemplateAnalyzer.static_datas)
                            except Exception as e:
                                print(f"错误：处理标签 {pd['text']} 时发生错误: {str(e)}")
                    else:
                        # 不是图片类型，处理所有标签
                        for pd in point_data:
                            label = TemplateAnalyzer.registered_labels[pd['type']]
                            try:
                                label.insert_data_to_point(pd, data, TemplateAnalyzer.static_datas)
                            except Exception as e:
                                print(f"错误：处理标签 {pd['text']} 时发生错误: {str(e)}")
                else:
                    # 单个标签的处理
                    label = TemplateAnalyzer.registered_labels[point_data['type']]
                    try:
                        label.insert_data_to_point(point_data, data, TemplateAnalyzer.static_datas)
                    except Exception as e:
                        print(f"错误：处理标签 {point_data['text']} 时发生错误: {str(e)}")
            except Exception as e:
                print(f"错误：处理标签 '{point_name}' 时发生未知错误: {str(e)}")

        return no_data_points
        
    @staticmethod
    def _smart_match_image_tags(image_tags, image_data, original_data):
        """智能匹配图片标签与图片数据
        
        Args:
            image_tags: 图片标签列表，每项是(标签名, 标签文本)的元组
            image_data: 图片数据字典，键是数据键名，值是图片数据元组
            original_data: 原始数据字典
            
        Returns:
            匹配结果字典，键是标签名，值是匹配到的图片数据
        """
        result = {}
        used_data = set()  # 记录已经使用过的数据
        
        # 遍历所有图片标签
        for tag_name, tag_text in image_tags:
            # 如果标签名已在原始数据中存在，且类型合适，则不需要额外匹配
            if tag_name in original_data:
                data = original_data[tag_name]
                if isinstance(data, tuple) and len(data) == 2 and isinstance(data[1], str):
                    if data[1].endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                        continue
            
            # 尝试多种匹配策略
            matched = False
            
            # 1. 直接匹配标签名
            if tag_name in image_data and tag_name not in used_data:
                result[tag_name] = image_data[tag_name]
                used_data.add(tag_name)
                matched = True
                continue
                
            # 2. 忽略大小写匹配
            tag_name_lower = tag_name.lower()
            for key, value in image_data.items():
                if key.lower() == tag_name_lower and key not in used_data:
                    result[tag_name] = value
                    used_data.add(key)
                    matched = True
                    break
            if matched:
                continue
                
            # 3. 检查"image:标签名"格式
            image_prefix_key = f"image:{tag_name}"
            for key, value in image_data.items():
                # 忽略大小写比较前缀
                if key.lower() == image_prefix_key.lower() and key not in used_data:
                    result[tag_name] = value
                    used_data.add(key)
                    matched = True
                    break
            if matched:
                continue
                
            # 4. 检查键名包含标签名的情况
            for key, value in image_data.items():
                if key not in used_data and (
                    tag_name.lower() in key.lower() or  # 键名包含标签名
                    key.lower().endswith(tag_name.lower())  # 键名以标签名结尾
                ):
                    result[tag_name] = value
                    used_data.add(key)
                    matched = True
                    break
            if matched:
                continue
                
            # 5. 检查图片路径包含标签名的情况
            for key, value in image_data.items():
                if key not in used_data and tag_name.lower() in value[1].lower():
                    result[tag_name] = value
                    used_data.add(key)
                    matched = True
                    break
            if matched:
                continue
                
            # 6. 如果都匹配不到，使用任意未使用的图片数据
            for key, value in image_data.items():
                if key not in used_data:
                    result[tag_name] = value
                    used_data.add(key)
                    matched = True
                    break
        
        return result

    @staticmethod
    def print_no_data_points(no_data_points):
        """打印无数据的插入点信息"""
        if len(no_data_points) == 0:
            return
        print("\n无数据对应的内容标签：")
        i = 1
        for point_name, point_data in no_data_points.items():
            if isinstance(point_data, list):
                for pd in point_data:
                    print(f"  ({i}) 标签名为'{point_name}'、类型为'{pd['type']}'的内容标签'{pd['text']}'无法匹配到数据")
                    i += 1
            else:
                print(
                    f"  ({i}) 标签名为'{point_name}'、类型为'{point_data['type']}'的内容标签'{point_data['text']}'无法匹配到数据")
                i += 1
