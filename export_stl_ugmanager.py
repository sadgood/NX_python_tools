# NX Journal: STL 导出（NX1899 API 兼容性 + Teamcenter 支持 + 装配导出）
#
import os
import sys
import traceback
import NXOpen
import NXOpen.UF

# --------------- 参数解析 ---------------
def parse_args():
    args = sys.argv[1:]

    if len(args) < 1:
        raise ValueError("用法: script.py <csv file> [out stl path] [ChordalTol] [AdjTol]")

    csv_file = args[0].strip()

    if len(args) >= 2 and args[1].strip():
        stl_path = os.path.abspath(args[1])

    chordal = 0.08
    adjacency = 0.08
    if len(args) >= 3 and args[2].strip():
        chordal = float(args[2])
    if len(args) >= 4 and args[3].strip():
        adjacency = float(args[3])

    return csv_file, stl_path, chordal, adjacency

# --------------- Teamcenter 工具函数 ---------------
def setup_teamcenter_config():
    """设置 Teamcenter 配置规则"""
    uf = NXOpen.UF.UFSession.GetUFSession()
    
    try:
        # 获取当前配置规则
        current_rule = uf.Ugmgr.AskConfigRule()
        print("[TC配置] 当前配置规则: %s" % current_rule)
        
        # 设置为 "Any Status; Working" 以获取最新版本
        new_rule = "Any Status; Working"
        uf.Ugmgr.SetConfigRule(new_rule)
        print("[TC配置] 设置配置规则: %s" % new_rule)
        
        return current_rule
        
    except Exception as e:
        print("[TC配置] 配置规则设置失败: %s" % e)
        return None

def restore_teamcenter_config(original_rule):
    """恢复原始 Teamcenter 配置规则"""
    if original_rule:
        try:
            uf = NXOpen.UF.UFSession.GetUFSession()
            uf.Ugmgr.SetConfigRule(original_rule)
            print("[TC配置] 恢复配置规则: %s" % original_rule)
        except Exception as e:
            print("[TC配置] 恢复配置规则失败: %s" % e)

def get_part_revision_info(item_id):
    """获取零件修订信息"""
    uf = NXOpen.UF.UFSession.GetUFSession()
    
    try:
        print("[TC查询] 查询零件: %s" % item_id)
        
        # 获取零件标签
        item_tag = uf.Ugmgr.AskPartTag(item_id)
        if not item_tag:
            raise RuntimeError("未找到零件: %s" % item_id)
        
        print("[TC查询] 零件标签: %s" % item_tag)
        
        # 获取配置的修订版本
        itemrev_tag = uf.Ugmgr.AskConfiguredRev(item_tag)
        if not itemrev_tag:
            raise RuntimeError("未找到配置的修订版本: %s" % item_id)
        
        print("[TC查询] 修订标签: %s" % itemrev_tag)
        
        # 获取修订 ID
        revision_id = uf.Ugmgr.AskPartRevisionId(itemrev_tag)
        print("[TC查询] 配置的修订 ID: %s" % revision_id)
        
        return {
            'item_tag': item_tag,
            'itemrev_tag': itemrev_tag,
            'revision_id': revision_id,
            'configured_revision': revision_id
        }
        
    except Exception as e:
        print("[TC查询] 查询失败: %s" % e)
        raise

def encode_part_filename(item_id, revision):
    """编码零件文件名"""
    uf = NXOpen.UF.UFSession.GetUFSession()
    
    try:
        # 方法1: 使用 EncodePartFilename
        encoded_name = uf.Ugmgr.EncodePartFilename(item_id, revision, "master", "")
        print("[TC编码] 编码文件名: %s" % encoded_name)
        return encoded_name
        
    except Exception as e:
        print("[TC编码] 编码失败，使用默认格式: %s" % e)
        # 方法2: 使用标准格式
        return "@DB/%s/%s" % (item_id, revision)

# --------------- 零件打开函数 ---------------
def open_part_by_item_id(session, item_id):
    """根据 item ID 和修订版本打开零件"""
    original_config = None
    
    try:
        # 设置 Teamcenter 配置
        original_config = setup_teamcenter_config()
        
        # 获取零件修订信息
        part_info = get_part_revision_info(item_id)
        final_revision = part_info['revision_id']
        
        print("[打开零件] 使用修订: %s" % final_revision)
        
        work_part = None
        
        try:
            # 检查是否已经打开
            for part in session.Parts:
                if hasattr(part, 'Name') and item_id in part.Name:
                    print("[方法1] 零件已在会话中，设为工作零件: %s" % part.Name)
                    session.Parts.SetWork(part)
                    return part, final_revision
                
            encoded_name = encode_part_filename(item_id, final_revision)
            
            session.Parts.LoadOptions.PartLoadOption = NXOpen.LoadOptions.LoadOption.FullyLoad
            session.Parts.LoadOptions.ComponentsToLoad = NXOpen.LoadOptions.LoadComponents.LastSet
            session.Parts.LoadOptions.UseLightweightRepresentations =  True
            session.Parts.LoadOptions.UsePartialLoading = False

            work_part, load_status = session.Parts.OpenBaseDisplay(encoded_name)
            if load_status:
                load_status.Dispose()
            
            if work_part:
                session.Parts.SetWork(work_part)
                return work_part, final_revision
                
        except Exception as e:
            print("打开失败: %s" % e)
            raise RuntimeError(f"{item_id} 打开失败，原因：{str(e)}")
                    
    except Exception as e:
        print("[打开零件] 失败: %s" % e)
        raise
        
    finally:
        # 恢复原始配置
        restore_teamcenter_config(original_config)

# --------------- 路径 ---------------
def ensure_output_directory(file_path):
    """确保输出目录存在"""
    directory = os.path.dirname(file_path)
    if directory and not os.path.exists(directory):
        try:
            os.makedirs(directory, exist_ok=True)
            print(f"创建输出目录: {directory}")
            return True
        except Exception as e:
            print(f"创建目录失败: {e}")
            return False
    return True

# --------------- 读取 CSV 文件 ---------------
def read_itemids_from_csv(csv_file):
    """从 CSV 文件读取 itemid 列表"""
    itemids = []
    try:
        with open(csv_file, 'r') as f:
            for line_num, line in enumerate(f, 1):
                line = line.strip()
                if line and not line.startswith('#'):
                    itemid = line.split(',')[0].strip()
                    if itemid:
                        itemids.append(itemid)
        print("[INFO] Read " + str(len(itemids)) + " itemids from CSV")
        return itemids
    except Exception as e:
        raise RuntimeError("Failed to read CSV file: " + str(e))
    
# --------------- 收集和分析体 ---------------
def analyze_bodies(work_part):
    body_list = []
    
    print("[详细诊断] 开始分析所有 Body...")
    
    # 转换 Bodies 集合为列表
    try:
        bodies_collection = work_part.Bodies
        if not bodies_collection:
            print("[信息] Bodies 集合为空")
            return  body_list
            
        for b in bodies_collection:
            body_list.append(b)
            
        print("[信息] 找到 %d 个 Body" % len(body_list))
        
    except Exception as e:
        print("[错误] 无法遍历 Bodies 集合: %s" % e)
        return  body_list
    
    return body_list

# --------------- 装配组件收集---------------
def try_collect_assembly_components(work_part):
    """收集装配组件并提取几何体"""
    components = []
    all_bodies = []
    
    try:           
        root_comp = work_part.ComponentAssembly.RootComponent
        if root_comp:
            def collect_recursive(comp):
                components.append(comp)
                try:
                    if hasattr(comp, "Prototype"):
                        prototype = comp.Prototype
                        if prototype and hasattr(prototype, "Bodies"):
                            proto_bodies = prototype.Bodies
                            if proto_bodies:
                                for body in proto_bodies:
                                    com_body = comp.FindOccurrence(body)
                                    all_bodies.append(com_body)

                except Exception as e:
                    print("    Prototype 方法失败: %s" % e)

                try:
                    for child in comp.GetChildren():
                        collect_recursive(child)
                except:
                    pass
            
            collect_recursive(root_comp)
            print("[装配信息] 找到组件数量: %d" % len(components))
            
        else:
            print("[装配信息] 无根组件，不是装配文件")
    except Exception as e:
        print("[装配信息] 装配访问异常: %s" % e)
    
    return components, all_bodies

# --------------- 导出函数 ---------------
def export_stl(session, objects_to_export, out_file, chordal_tol, adjacency_tol):
    if not objects_to_export:
        raise RuntimeError("导出对象列表为空")
    
    print("[导出] 准备导出 %d 个对象到: %s" % (len(objects_to_export), out_file))
    
    # 验证导出对象类型
    valid_objects = []
    for i, obj in enumerate(objects_to_export):
        try:
            # 检查对象类型
            obj_type = type(obj).__name__
            print("  对象[%d] 类型: %s" % (i, obj_type))
            
            # 只接受 Body 类型的对象
            if "Body" in obj_type or hasattr(obj, "GetFaces"):
                valid_objects.append(obj)
                print("    ✓ 有效的几何体对象")
            else:
                print("    ❌ 跳过非几何体对象")
        except Exception as e:
            print("    对象验证异常: %s" % e)
    
    if not valid_objects:
        raise RuntimeError("没有有效的几何体对象可以导出")
    
    print("[导出] 有效对象数量: %d" % len(valid_objects))
    
    stl_creator = session.DexManager.CreateStlCreator()
    try:
        stl_creator.AutoNormalGen = True
        stl_creator.ChordalTol = chordal_tol
        stl_creator.AdjacencyTol = adjacency_tol
        stl_creator.TriangleDisplay = True
        
        # 添加到选择集
        stl_creator.ExportSelectionBlock.Add(valid_objects)
        stl_creator.OutputFile = out_file
        
        # 执行导出
        mark_id = session.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Export STL")
        try:
            stl_obj = stl_creator.Commit()
            session.DeleteUndoMark(mark_id, None)
            print("[导出] STL 导出成功")
            return stl_obj
        except Exception as e:
            session.UndoToMark(mark_id, None)
            raise RuntimeError("STL 导出失败: %s" % e)
    finally:
        try:
            stl_creator.Destroy()
        except:
            pass

# --------------- 主函数 ---------------
def main():
    the_session = NXOpen.Session.GetSession()
    work_part = None
    
    try:
        csv_file, stl_path, chordal, adjacency = parse_args()
        print("[信息] 零件 ID: %s" % csv_file)
        print("[信息] 输出 STL: %s" % stl_path)
        print("[信息] 容差设置: ChordalTol=%.6f AdjacencyTol=%.6f" % (chordal, adjacency))
        print()

        # 读取 itemid 列表
        itemids = read_itemids_from_csv(csv_file)
        if not itemids:
            raise RuntimeError("No valid itemids found in CSV file")
        
        print("[START] Batch processing " + str(len(itemids)) + " items...")
        print("=" * 60)

        if not ensure_output_directory(stl_path):
            print("无法创建STL输出目录，导出终止")
            exit()

        # 处理每个项目
        results = []
        success_count = 0
        
        for i, itemid in enumerate(itemids, 1):
            try:
                work_part, rev = open_part_by_item_id(the_session, itemid)

                #displayPart = the_session.Parts.Display
                #print("[尝试] displayPart ...%s"% displayPart)
                #print("[尝试] work_part ...%s"% work_part)

                root_comp = work_part.ComponentAssembly.RootComponent
                #print("[尝试] root_comp ...%s"% root_comp)
              
                objects_to_export = []
            
                if  root_comp:
                    print("[收集] 检查装配组件...")
                    components, assembly_bodies = try_collect_assembly_components(work_part)
                    if assembly_bodies:
                        objects_to_export = assembly_bodies
                        print("[选择] 使用从装配中提取的 %d 个几何体" % len(assembly_bodies))           
                else:
                    body_list = None
                    body_list = analyze_bodies(work_part)
                    if body_list:
                        objects_to_export= body_list

                if objects_to_export:
                    success_count += 1
                    out_stl = os.path.join(stl_path, itemid +"_"+ rev + ".stl")
                    export_stl(the_session, objects_to_export, out_stl, chordal, adjacency)
                else:
                    results.append(f"{itemid} 可导出体：{len(objects_to_export)}")
                
                resp = NXOpen.PartCloseResponses()
                the_session.Parts.CloseAll(NXOpen.BasePartCloseModified.CloseModified, resp)

            except Exception as e:
                results.append(f"{itemid} 导stl报错：{str(e)}")

        # 保存报告
        report_file = os.path.join(stl_path, "export_report.txt")
        try:
            with open(report_file, 'w') as f:
                f.write("STL Export Report\n")
                f.write("=" * 50 + "\n")
                f.write("Total: " + str(len(itemids)) + ", Success: " + str(success_count) + ", Failed: " + str(len(itemids) - success_count) + "\n")
                if itemids:
                    success_rate = 100.0 * success_count / len(itemids)
                    f.write("Success rate: " + str(round(success_rate, 1)) + "%\n\n")
                
                f.write("Detailed Results:\n")
                f.write("-" * 50 + "\n")
                for result in results:
                    f.write(result)
                    f.write("\n")
            print("[REPORT] Detailed report saved: " + report_file)
        except Exception as e:
            print("[WARNING] Cannot save report: " + str(e))
            
    except Exception as e:
        print("[错误] 导出失败: %s" % e)
        traceback.print_exc()
        raise

if __name__ == "__main__":
    main()