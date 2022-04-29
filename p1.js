//resource used:https://www.youtube.com/watch?v=R9NbXPuT0ro





function heldkarp(graph){

    var start= performance.now();

    //memo={};

    //create places to store our cities and distances
    mycities=Array.from(Array(graph.length).keys());
    distances=new Array(graph.length);
    //generate the shortest path starting and ending at each vertex and store them
    for (let i=0; i<mycities.length;i++){
        console.log('memo cleared');
        memo={};
        distances[i]=heldkarpex(mycities, i, graph);
    }
    //return the minimum of those distances
    result =distances[0];
    for(let j=1;j<distances.length;j++){
        if (distances[j]<result){
            result=distances[j];
        }
    }
    var end=performance.now();
    console.log(graph.length, " : " ,end-start);
    return result;
}



function heldkarpex(cities, start ,graph){
    //if only two cities left, we return that distance between them
    if(cities.length==2){
        //find what the other vertex is
        var other=cities[0];
        if(other==start){
            other=cities[1];
        }
        d=graph[start][other];
        console.log('added');
        memo[[cities,start]]=d;
        return d;
    }

    else{
        //remove start from cities
        var temp=[...cities];
        var index = temp.indexOf(start);    
        temp.splice(index, 1);

        //find and return the minimum of the recursive calls
        var min =Infinity;
        for (let i=0;i<cities.length;i++){
            
            if(cities[i]!=start){
                test=temp.concat(cities[i]);
                m=memo[test];
                if(memo[test]!=undefined){
                    //console.log('skipped');
                    p=memo[test]+graph[start][cities[i]];
                }
                else{
                    p=heldkarpex(temp,cities[i],graph)+graph[start][cities[i]];
                    
                }
                if(p<min){
                    min=p;
                }
            }
        }
        console.log('added');
        memo[[cities,start]]=min;
        return min;
        
    }
}


function generateAdjMat(d){
    
    var res=Array.from(Array(d), () => new Array(d).fill(0));
    for (let i=0;i<d;i++){
        for(let k=0;k<i;k++){
            res[i][k]=Math.floor(Math.random()*9)+1;
            res[k][i]=res[i][k];        
        }
    }
    return res
}

function generateAdjMatAllOnes(d){
    
    var res=Array.from(Array(d), () => new Array(d).fill(0));
    for (let i=0;i<d;i++){
        for(let k=0;k<i;k++){
            res[i][k]=1;
            res[k][i]=res[i][k];        
        }
    }
    return res
}
/*
console.log(generateAdjMat(10));

graph1=[[0,1,2],
        [1,0,2],
        [2,2,0]];

graph2=[[0,1,2,2],
        [1,0,2,2],
        [2,2,0,2],
        [2,2,2,0]];

graph3=generateAdjMatAllOnes(10);
graph4=generateAdjMat(10);


dist1=heldkarp(graph1);
console.log(dist1);
dist2=heldkarp(graph2);
console.log(dist2);
dist3=heldkarp(graph3);

console.log(heldkarp(graph3));
console.log(heldkarp(graph4));
*/

for(let i=10;i<101;i++){
    myMat=generateAdjMatAllOnes(i);
    console.log(heldkarp(myMat));
}