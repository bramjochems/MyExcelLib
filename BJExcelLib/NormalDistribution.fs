namespace BJExcelLib.Math

/// Module with functions for the normal distribution (univariate and bivariate at the moment)
module public NormalDistribution =

    // Other internal helper functions, self explanatory what they do
    let private dist = MathNet.Numerics.Distributions.Normal()
    let public nd x = dist.Density x
    let public  cnd x = dist.CumulativeDistribution x
    let public  cndc x = 1. - cnd x


    /// Density of a bivariate normal distribution where the first marginal distribution is N(m1,s1^2) distributed and the second
    /// marginal is N(m2,s2^2). The correlation between the variables is rho. The distribution is evaluated at X1=x1 and X2=x2.
    let public pdf2 (m1,s1) (m2,s2) rho x1 x2 =
        let x1 = (x1-m1)/s1
        let x2 = (x2-m2)/s2
        let rfac = (1.-rho*rho)
        exp(-0.5*(x1*x1-2.*rho*x1*x2+x2*x2)/rfac)/(rfac*s1*s2)
 
 
    /// Some constants for the below bivariate cumulative distribution function
    let private weights1 = [ 0.17132449237917; 0.17132449237917;0.360761573048138; 0.360761573048138;0.46791393457269;0.46791393457269 ]
    let private values1 =  [ 0.966234757101576;0.033765242898424;0.830604693233133; 0.169395306766867;0.619309593041598;0.380690406958402 ]
    let private weights2 = [ 0.0471753363865118; 0.0471753363865118; 0.106939325995318; 0.106939325995318;0.160078328543346;  0.160078328543346;  0.203167426723066; 0.20316742672306; 0.233492536538355;  0.233492536538355;  0.249147045813403; 0.249147045813403]
    let private values2 =  [ 0.99078031712336; 0.00921968287664;0.95205862818524; 0.04794137181476;0.88495133709715;0.11504866290285;0.79365897714331;0.20634102285669;0.68391574949909;0.31608425050091;0.56261670425574;0.43738329574427]
    let private weights3 = [0.0176140071391521; 0.0176140071391521; 0.0406014298003869;0.0406014298003869; 0.0626720483341091; 0.0626720483341091;0.0832767415767048; 0.0832767415767048; 0.1019301198172400;0.1019301198172400; 0.1181945319615180; 0.1181945319615180;0.1316886384491770; 0.1316886384491770; 0.1420961093183820;0.1420961093183820; 0.1491729864726040; 0.1491729864726040;0.1527533871307260; 0.1527533871307260]
    let private values3 = [ 0.99656429959255; 0.00343570040745; 0.98198596363896;0.01801403636104; 0.95611721412566; 0.04388278587434;0.91955848591111; 0.08044151408889; 0.87316595323008;0.12683404676992; 0.81802684036326; 0.18197315963674;0.75543350097541; 0.24456649902459; 0.68685304435771;0.31314695564229; 0.61389292557082; 0.38610707442918;0.53826326056675; 0.46173673943325 ]
    let private values4 = [ 0.000047216149159;3.972561612889520;0.001298022024068;3.857185731135710;0.007702795584372;3.656640508589690;0.025883348755653;3.382351236044530;0.064296126805960;3.050028889390040;0.136503050785829;2.658650279846430;0.239251089780572;2.282719097583890;0.392244063312137;1.887068418173820;0.596314691697034;1.507458096263630;0.984753258857668;1.015363867311070 ]

 
    /// Cumulative bivariate normal distribution, adapted from Haug's Option Pricing formulas,
    /// who credits Graeme West, who implemented an algorithm of Genze. (mi,si) are the mean and
    /// standard deviation of the i-th variable, rho is the correlation and x1 and x2 are the
    /// points where the cumulative distribution is evaluated at.
    let public cdf2 (m1,s1) (m2,s2) rho x1 x2 =
        let x = (x1-m1)/s1
        let y = (x2-m2)/s2
 
        let rho = min 1. (max -1. rho)
        let h = -x
        let k = if rho < -0.925 then y else -y
        let hk = h * k
        match rho with
            | x when x = 0.         ->  cnd(-h) * cnd(-k)
            | x when rho = 1.       ->  cnd(-(max h k))
            | x when rho = -1.      ->  if k > h then cnd(k)-cnd(h) else 0.
            | x when abs(x) < 0.925 ->  let weights, values = match x with
                                                                | r when abs(r) < 0.3    -> weights1,values1
                                                                | r when abs(r) < 0.75   -> weights2,values2
                                                                | _                      -> weights3,values3
                                        let hs = 0.5*(h*h+k*k)
                                        let asr_ = System.Math.Asin(rho)
                                        values  |> List.map(fun z -> let sn = System.Math.Sin(asr_ * z)
                                                                     exp((sn*hk-hs)/(1.-sn*sn)) )
                                                |> List.zip weights
                                                |> List.sumBy (fun (w,v) -> w*v)
                                                |> fun x -> x * asr_/(4.*System.Math.PI) + cnd(-h) *cnd(-k)
 
            | _                     ->  // 0.925 < abs(x) < 1
                                        let ass = (1.-rho)*(1.+rho)
                                        let A = sqrt(ass)
                                        let bs = (h-k)*(h-k)
                                        let c = (4.-hk) / 8.
                                        let d = (12.-hk)/16.
                                        let asr_ = -0.5*(bs / ass + hk)
                                        let BVN = if asr_ <= -100. then 0.
                                                  else A*exp(asr_)*(1.-c*(bs-ass)*(1.-0.2*d*bs) / 3. + 0.2*c* d*ass*ass)
                                        let BVN = if -hk >= 100. then BVN
                                                  else
                                                    let B = sqrt(bs)
                                                    BVN - sqrt(2.*System.Math.PI)*exp(-0.5*hk)*cnd(-B/A)*B*(1.-c*bs*(1.-0.2*d*bs)/3.)    
 
                                        let A = 0.5*A
                                        let A2 = A*A
 
                                        values4 |> List.zip weights3
                                                |> List.sumBy(fun (w,x) ->  let xs = A2*x
                                                                            let asr_ = -0.5*(bs / xs + hk)
                                                                            if asr_ < -100. then 0.
                                                                            else
                                                                                let rs = sqrt(1.-xs)
                                                                                A*w*exp(asr_)*(exp(-hk*(1.-rs)/(2.*(1.+rs)))/rs - (1.+c*xs*(1.+d*xs))))
                                                |> fun x -> let BVN = -(BVN+x)/(2.*System.Math.PI)
                                                            match rho > 0. with
                                                                | true  ->  BVN + cnd(-(max h k))
                                                                | false ->  if k > h then -BVN + cnd(k)-cnd(h) else -BVN

