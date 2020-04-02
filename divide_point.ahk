; 2020-04-01 , 두점의 m 과 n 의 내분점  좌표 
divide_inside_point_gap(point_1 , point_2 , m , n , val_round := 4)
{
    ; point_1, point_2 - [ x pos , y pos ]
    temp 	:= []          ; 반환 배열
    temp[1] := Round( (m*point_2[1] + n*point_1[1]) / (m + n) , val_round )
    temp[2] := Round( (m*point_2[2] + n*point_1[2]) / (m + n) , val_round )
    temp[3] := abs(temp[1] - point_1[1]) ; point_1과의 x 거리 값
    temp[4] := abs(temp[2] - point_1[2]) ; point_1과의 y 거리 값
    temp[5] := abs(temp[1] - point_2[1]) ; point_2과의 y 거리 값
    temp[6] := abs(temp[2] - point_2[2]) ; point_2과의 y 거리 값
    
    return temp
}
; 2020-04-01 , 두점의 m 과 n 의 외분점  좌표 
divide_outside_point_gap(point_1 , point_2 , m , n , val_round := 4)
{
    ; point_1, point_2 - [ x pos , y pos ]
    temp 	:= []          ; 반환 배열
    temp[1] := Round( (m*point_2[1] - n*point_1[1]) / (m - n) , val_round )
    temp[2] := Round( (m*point_2[2] - n*point_1[2]) / (m - n) , val_round )
    temp[3] := abs(temp[1] - point_1[1]) ; point_1과의 x 거리 값
    temp[4] := abs(temp[2] - point_1[2]) ; point_1과의 y 거리 값
    temp[5] := abs(temp[1] - point_2[1]) ; point_2과의 y 거리 값
    temp[6] := abs(temp[2] - point_2[2]) ; point_2과의 y 거리 값
    
    return temp
}


