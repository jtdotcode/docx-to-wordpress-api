using Newtonsoft.Json;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace DocxEmailToWordPress
{
    class WpAPI
    {
        //public IEnumerable GetWordPressPosts()
        //{
        //    using (WebClient webClient = new WebClient())
        //    {
        //        string blogUrl = @"https://public-api.wordpress.com/rest/v1.1/sites/hendrikbulens.wordpress.com/posts/";
        //        string response = webClient.DownloadString(blogUrl);

        //        WordPressBlog blogPosts = JsonConvert.DeserializeObject(response);

        //        IEnumerable posts = blogPosts.posts.OrderByDescending(x => x.date).Take(3).Select(x => new BlogPost()
        //        {
        //            Title = x.title,
        //            Message = x.excerpt.CharacterLimit(150),
        //            PublishedOn = x.date,
        //            Comments = x.discussion.comment_count,
        //            Link = x.URL
        //        });

        //        return posts;
        //    }
        }


    }

