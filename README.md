# jsonToWord

Note: **This project aim to convert JSON file to word file**

Example:

input:

```
/**
*	demo.json
*/

{
    "info": {
      "title": "示例文档",
      "version": "1.0.0",
      "description": "这是一个示例文档"
    },
    "host": "https://localhost:3000",
    "basePath": "/v1",
    "interfaces": [
      {
        "key": "song",
        "title": "歌曲相关接口",
        "children": [{
            "key": "/song",
            "title": "获取歌曲列表",
            "path": "/song",
            "method": "get",
            "description": "获取接口分页列表",
            "parameters": [{
                "name": "page",
                "in": "query",
                "description": "页码",
                "required": true,
                "type": "number"
              },
              ,
              {
                "name": "pageSize",
                "in": "query",
                "description": "每页数据的条数",
                "required": true,
                "type": "number"
              }
            ],
            "responses": {
              "success": {
                "type": "object",
                "example": {
                  "pages": 3,
                  "count": 30,
                  "list": [{
                      "song_id": 7149583,
                      "name": "告白气球",
                      "url": "http://47.107.229.37:8081/gm/music/336/555949/gaobaiqiqiu.mp3",
                      "artist": "周杰伦",
                      "pic120": "http://img2.kuwo.cn/star/albumcover/120/64/39/3540704654.jpg",
                      "pic": "http://img2.kuwo.cn/star/albumcover/300/64/39/3540704654.jpg",
                      "album_name": "周杰伦的床边故事"
                    },
                    {
                      "song_id": 76323299,
                      "name": "说好不哭 (with 五月天阿信)",
                      "url": "http://47.107.229.37:8081/gm/music/336/10685968/shuohaobuku.mp3",
                      "artist": "周杰伦",
                      "pic120": "http://img1.kuwo.cn/star/albumcover/120/12/37/4156270827.jpg",
                      "pic": "http://img1.kuwo.cn/star/albumcover/300/12/37/4156270827.jpg",
                      "album_name": "说好不哭 (with 五月天阿信)"
                    }
                  ]
                }
              }
            },
            "response_description": [
              {
                "name": "song_id",
                "description": "歌曲id"
              },
              {
                "name": "name",
                "description": "歌曲名称"
              },
              {
                "name": "url",
                "description": "歌曲链接"
              },
              {
                "name": "artist",
                "description": "演唱者"
              },
              {
                "name": "pic",
                "description": "歌曲封面图片"
              },
              {
                "name": "album_name",
                "description": "专辑名称"
              }
            ] 
          },
          ...
```



output:

## 示例文档

v1.0.0
Base URL: https://localhost:3000/v1
这是一个示例文档

### 获取歌曲列表
#### URL:
​    GET  /song
#### Parameter:
​    page: <number> 页码
​    pageSize: <number> 每页数据的条数
#### Response:
{
  "pages": 3,
  "count": 30,
  "list": [
    {
      "song_id": 7149583,
      "name": "告白气球",
      "url": "http://47.107.229.37:8081/gm/music/336/555949/gaobaiqiqiu.mp3",
      "artist": "周杰伦",
      "pic120": "http://img2.kuwo.cn/star/albumcover/120/64/39/3540704654.jpg",
      "pic": "http://img2.kuwo.cn/star/albumcover/300/64/39/3540704654.jpg",
      "album_name": "周杰伦的床边故事"
    },
    {
      "song_id": 76323299,
      "name": "说好不哭 (with 五月天阿信)",
      "url": "http://47.107.229.37:8081/gm/music/336/10685968/shuohaobuku.mp3",
      "artist": "周杰伦",
      "pic120": "http://img1.kuwo.cn/star/albumcover/120/12/37/4156270827.jpg",
      "pic": "http://img1.kuwo.cn/star/albumcover/300/12/37/4156270827.jpg",
      "album_name": "说好不哭 (with 五月天阿信)"
    }
  ]
}
#### Response_Description:
​    song_id: 歌曲id
​    name: 歌曲名称
​    url: 歌曲链接
​    artist: 演唱者
​    pic: 歌曲封面图片
​    album_name: 专辑名称

