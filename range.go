package goexcelerate


type Range struct {
	Start string
	End string
	IsRow bool
	IsColumn bool
	IsCell bool
	X int
	Y int
	Height int
	Width int
}

func NewRange(rng string) {
	
}