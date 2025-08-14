 

def outlook_depth_pathfinder(target_folder_l,n):
    # working outlook folder probe. finds path of deepest folder(s)
    # #always make n=1. represents start layer. increments deeper func goes
    last_count=0
    target_l=[]
    for target_folder in target_folder_l:
        if not eval(target_folder):
            last_count+=1
            if last_count==len(target_folder_l):
                return target_folder_l # return list of all deepest folders
            else:
                pass
        else:
            n=n+1
            f_temp="['{}'].Folders"
            for folder in eval(target_folder): 
                new_f=f_temp.format(str(folder.Name))
                target_f=target_folder+new_f
                target_l.append(target_f)
    next_layer=outlook_depth_pathfinder(target_l,n) 
    return next_layer


def outlook_depth_detector(target_folder_l,n):
    # depth of deepest subdirectory of given folder. always make n=1. represents start layer. increments deeper func goes
    last_count=0
    target_l=[]
    for target_folder in target_folder_l:
        if not eval(target_folder):
            last_count+=1
            if last_count==len(target_folder_l):
                report="Completed scanning:{f}"+"\n"+"Max depth:{n}"+"\n"+"Example of deepest folder:{e}"
                return report.format(f=target_folder_l[0].split(".Folders['",1)[0],
                                     n=n,
                                     e=target_folder_l[0]) # return max depth
            else:
                pass
        else:
            f_temp="['{}'].Folders"
            for folder in eval(target_folder): 
                new_f=f_temp.format(str(folder.Name))
                target_f=target_folder+new_f
                target_l.append(target_f)
    n=n+1
    next_layer=outlook_depth_detector(target_l,n) 
    return next_layer
